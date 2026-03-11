import json
import logging
import os
import re
import sys
import time
import threading
import ctypes
import ctypes.wintypes
from html.parser import HTMLParser
from pathlib import Path
from datetime import datetime

import pyperclip
import win32clipboard
from PIL import Image, ImageDraw, ImageFont
from pystray import Icon, Menu, MenuItem

from providers import load_provider_from_config

# ── HTML-to-Text Converter (preserves lists) ──────────────────────────
class _HTMLToText(HTMLParser):
    """Convert clipboard HTML to plain text preserving list formatting."""

    def __init__(self):
        super().__init__()
        self._parts = []
        self._in_ol = False
        self._in_ul = False
        self._ol_counter = 0

    def handle_starttag(self, tag, attrs):
        if tag == "ol":
            self._in_ol = True
            self._ol_counter = 0
        elif tag == "ul":
            self._in_ul = True
        elif tag == "li":
            if self._in_ol:
                self._ol_counter += 1
                self._parts.append(f"\n{self._ol_counter}. ")
            elif self._in_ul:
                self._parts.append("\n- ")
        elif tag in ("br", "p", "div"):
            if self._parts and not self._parts[-1].endswith("\n"):
                self._parts.append("\n")

    def handle_endtag(self, tag):
        if tag == "ol":
            self._in_ol = False
        elif tag == "ul":
            self._in_ul = False
        elif tag in ("p", "div", "li"):
            if self._parts and not self._parts[-1].endswith("\n"):
                self._parts.append("\n")

    def handle_data(self, data):
        self._parts.append(data)

    def get_text(self):
        return "".join(self._parts).strip()


def get_clipboard_text():
    """Read clipboard, preferring HTML format to preserve list formatting."""
    try:
        win32clipboard.OpenClipboard()
        try:
            cf_html = win32clipboard.RegisterClipboardFormat("HTML Format")
            if win32clipboard.IsClipboardFormatAvailable(cf_html):
                html_bytes = win32clipboard.GetClipboardData(cf_html)
                html_str = html_bytes.decode("utf-8", errors="replace")
                start = html_str.find("<!--StartFragment-->")
                end = html_str.find("<!--EndFragment-->")
                if start != -1 and end != -1:
                    fragment = html_str[start + 20:end]
                    if re.search(r"<[ou]l|<li", fragment, re.IGNORECASE):
                        parser = _HTMLToText()
                        parser.feed(fragment)
                        result = parser.get_text()
                        if result:
                            return result
        finally:
            win32clipboard.CloseClipboard()
    except Exception:
        pass
    return pyperclip.paste()


# ── Config ──────────────────────────────────────────────────────────────
MIN_WORDS = 5
MAX_WORDS = 500

# Use AppData for data files (reliable path for both .py and .exe)
if getattr(sys, "frozen", False):
    # Running as PyInstaller exe
    APP_DIR = Path(os.environ.get("LOCALAPPDATA", "")) / "ClipboardCoach"
else:
    # Running as Python script
    APP_DIR = Path(__file__).parent

APP_DIR.mkdir(parents=True, exist_ok=True)
HISTORY_FILE = APP_DIR / "coaching-history.json"
LOG_FILE = APP_DIR / "clipfix.log"

BACKGROUND_MODE = "--background" in sys.argv

# ── Logging (always log to file + console when available) ──────────────
log = logging.getLogger("coach")
log.setLevel(logging.INFO)

# Always write to log file
file_handler = logging.FileHandler(str(LOG_FILE), encoding="utf-8")
file_handler.setFormatter(logging.Formatter("%(asctime)s %(message)s", datefmt="%Y-%m-%d %H:%M:%S"))
log.addHandler(file_handler)

# Also log to console if not background mode
if not BACKGROUND_MODE:
    console_handler = logging.StreamHandler()
    console_handler.setFormatter(logging.Formatter("%(message)s"))
    log.addHandler(console_handler)


# ── Notification (via system tray balloon) ─────────────────────────────
def silent_notify(title, line2, line3=None):
    """Show a notification via the system tray icon balloon."""
    msg = line2
    if line3:
        msg = f"{line2}\n{line3}"
    try:
        if tray_icon and tray_icon.visible:
            tray_icon.notify(msg, title)
        else:
            log.info("  [notify] (tray not ready) %s: %s", title, msg)
    except Exception as e:
        log.warning("Notification failed: %s", e)


# ── LLM Provider (loaded at startup) ──────────────────────────────────
provider = None


# ── Pattern Tracking ────────────────────────────────────────────────────
history = []
if HISTORY_FILE.exists():
    try:
        history = json.loads(HISTORY_FILE.read_text())
    except Exception:
        history = []


def save_history():
    trimmed = history[-100:]
    HISTORY_FILE.write_text(json.dumps(trimmed, indent=2))


def get_top_patterns():
    counts = {}
    for entry in history:
        issue = entry.get("issue")
        if issue:
            counts[issue] = counts.get(issue, 0) + 1
    return [
        f"{issue} ({count}x)"
        for issue, count in sorted(counts.items(), key=lambda x: -x[1])[:3]
    ]


# ── Filters ─────────────────────────────────────────────────────────────
def looks_like_message(text):
    trimmed = text.strip()

    word_count = len(trimmed.split())
    if word_count < MIN_WORDS or word_count > MAX_WORDS:
        return False

    if re.match(r"^[\s]*[{(\[]", trimmed):
        return False
    code_chars = len(re.findall(r"[{}();=<>]", trimmed))
    if code_chars > len(trimmed) * 0.05:
        return False
    if re.match(r"^(import |from |const |let |var |function |class |def |#include)", trimmed):
        return False

    if re.match(r"^https?://", trimmed):
        return False
    if re.match(r"^[A-Z]:\\", trimmed) or re.match(r"^/[a-z]", trimmed):
        return False

    if " " not in trimmed and "\t" not in trimmed:
        return False

    return bool(re.search(
        r"\b(I|you|we|the|this|that|please|can|will|would|should|could|"
        r"hi|hey|hello|thanks|thank|let|just|want|need|think|know|like|"
        r"get|make|do)\b",
        trimmed,
        re.IGNORECASE,
    ))


# ── Analysis Cache ─────────────────────────────────────────────────────
_cache = {}
MAX_CACHE = 50


# ── Analysis via LLM Provider ─────────────────────────────────────────
def analyze_message(text):
    cache_key = text.strip()
    if cache_key in _cache:
        log.info("  (cached result)")
        return _cache[cache_key], 0.0

    top_patterns = get_top_patterns()
    pattern_hint = ""
    if top_patterns:
        pattern_hint = f"\nWatch for: {', '.join(top_patterns)}"

    result, api_duration = provider.analyze(text, pattern_hint)

    if len(_cache) >= MAX_CACHE:
        _cache.pop(next(iter(_cache)))
    _cache[cache_key] = result

    return result, api_duration


# ── Rewrite State ──────────────────────────────────────────────────────
pending_rewrite = {"current": None, "pasted": False}


def _simulate_paste():
    """Simulate Ctrl+V using Win32 keybd_event."""
    VK_CONTROL = 0x11
    VK_V = 0x56
    KEYEVENTF_KEYUP = 0x0002
    user32.keybd_event(VK_CONTROL, 0, 0, 0)
    user32.keybd_event(VK_V, 0, 0, 0)
    user32.keybd_event(VK_V, 0, KEYEVENTF_KEYUP, 0)
    user32.keybd_event(VK_CONTROL, 0, KEYEVENTF_KEYUP, 0)


def on_ctrl_m():
    """Global hotkey: Ctrl+M pastes the rewrite into the active window."""
    if pending_rewrite["current"] and not pending_rewrite["pasted"]:
        pyperclip.copy(pending_rewrite["current"])
        time.sleep(0.05)
        _simulate_paste()
        pending_rewrite["pasted"] = True
        log.info("  [OK] Rewrite pasted via Ctrl+M!")
        silent_notify("Clipboard Coach", "Rewrite pasted!")


# ── Display ─────────────────────────────────────────────────────────────
def display_result(result):
    if result["verdict"] == "good":
        log.info("[OK] Looks good -- send it.")
        silent_notify("Clipboard Coach", "Looks good -- send it!")
        return None

    issue = result.get("issue", "")
    nudge = result.get("nudge", "")
    rewrite = result.get("rewrite")

    log.info(">> %s", issue.upper() if issue else "SUGGESTION")
    log.info("   %s", nudge)
    if rewrite:
        log.info("   Rewrite: %s", rewrite)

    if rewrite and len(rewrite) < 200:
        silent_notify(
            issue or "Communication Coach",
            nudge,
            f"Rewrite: {rewrite}  --  Ctrl+M to paste",
        )
    elif rewrite:
        silent_notify(issue or "Communication Coach", nudge, "Ctrl+M to paste rewrite")
    else:
        silent_notify(issue or "Communication Coach", nudge)

    if issue:
        history.append({"issue": issue, "timestamp": datetime.now().isoformat()})
        save_history()

    return rewrite


# ── Background Analysis ────────────────────────────────────────────────
analyzing_lock = threading.Lock()


def analyze_in_background(text, t_detected):
    """Run analysis in a background thread to keep polling responsive."""
    def _run():
        try:
            result, api_duration = analyze_message(text)
            rewrite = display_result(result)
            t_total = time.perf_counter() - t_detected
            log.info("  [timing] End-to-end: %.2fs (API: %.2fs, overhead: %.2fs)",
                     t_total, api_duration, t_total - api_duration)
            if rewrite:
                pending_rewrite["current"] = rewrite
                pending_rewrite["pasted"] = False
                log.info("  -> Press Ctrl+M anywhere to paste the rewrite")
        except Exception as e:
            log.error("[!] Error: %s", e)
        finally:
            analyzing_lock.release()

    if analyzing_lock.acquire(blocking=False):
        log.info("Detected message -- analyzing...")
        threading.Thread(target=_run, daemon=True).start()


# ── Windows Clipboard Listener + Hotkey ────────────────────────────────
WM_CLIPBOARDUPDATE = 0x031D
WM_HOTKEY = 0x0312
WM_DESTROY = 0x0002
MOD_CONTROL = 0x0002
VK_M = 0x4D
HOTKEY_ID_CTRL_M = 1

user32 = ctypes.windll.user32
kernel32 = ctypes.windll.kernel32

WNDPROC = ctypes.WINFUNCTYPE(
    ctypes.c_longlong,
    ctypes.c_void_p,
    ctypes.c_uint,
    ctypes.c_longlong,
    ctypes.c_longlong,
)


class WNDCLASSW(ctypes.Structure):
    _fields_ = [
        ("style", ctypes.wintypes.UINT),
        ("lpfnWndProc", ctypes.c_void_p),
        ("cbClsExtra", ctypes.c_int),
        ("cbWndExtra", ctypes.c_int),
        ("hInstance", ctypes.wintypes.HINSTANCE),
        ("hIcon", ctypes.wintypes.HICON),
        ("hCursor", ctypes.wintypes.HANDLE),
        ("hbrBackground", ctypes.wintypes.HBRUSH),
        ("lpszMenuName", ctypes.wintypes.LPCWSTR),
        ("lpszClassName", ctypes.wintypes.LPCWSTR),
    ]


user32.CreateWindowExW.argtypes = [
    ctypes.wintypes.DWORD, ctypes.wintypes.LPCWSTR, ctypes.wintypes.LPCWSTR,
    ctypes.wintypes.DWORD,
    ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_int,
    ctypes.wintypes.HWND, ctypes.wintypes.HMENU,
    ctypes.wintypes.HINSTANCE, ctypes.wintypes.LPVOID,
]
user32.CreateWindowExW.restype = ctypes.wintypes.HWND
user32.DefWindowProcW.argtypes = [ctypes.c_void_p, ctypes.c_uint, ctypes.c_longlong, ctypes.c_longlong]
user32.DefWindowProcW.restype = ctypes.c_longlong


def create_clipboard_listener(callback):
    """Create a hidden window that listens for clipboard changes."""
    last_text = {"value": ""}

    def wnd_proc(hwnd, msg, wparam, lparam):
        if msg == WM_CLIPBOARDUPDATE:
            try:
                current_text = get_clipboard_text()
                if current_text != last_text["value"]:
                    last_text["value"] = current_text
                    if looks_like_message(current_text):
                        callback(current_text, time.perf_counter())
            except Exception as e:
                log.error("[!] Clipboard read error: %s", e)
            return 0
        elif msg == WM_HOTKEY and wparam == HOTKEY_ID_CTRL_M:
            on_ctrl_m()
            return 0
        elif msg == WM_DESTROY:
            user32.PostQuitMessage(0)
            return 0
        return user32.DefWindowProcW(hwnd, msg, wparam, lparam)

    wnd_proc_cb = WNDPROC(wnd_proc)

    wc = WNDCLASSW()
    wc.lpfnWndProc = ctypes.cast(wnd_proc_cb, ctypes.c_void_p)
    wc.hInstance = kernel32.GetModuleHandleW(None)
    wc.lpszClassName = "ClipboardCoachListener"

    class_atom = user32.RegisterClassW(ctypes.byref(wc))
    if not class_atom:
        raise RuntimeError("Failed to register window class")

    hwnd = user32.CreateWindowExW(
        0, wc.lpszClassName, "Clipboard Coach",
        0, 0, 0, 0, 0,
        None, None, wc.hInstance, None,
    )
    if not hwnd:
        raise RuntimeError("Failed to create listener window")

    if not user32.AddClipboardFormatListener(hwnd):
        raise RuntimeError("Failed to add clipboard listener")

    # Register Ctrl+M global hotkey
    if user32.RegisterHotKey(hwnd, HOTKEY_ID_CTRL_M, MOD_CONTROL, VK_M):
        log.info("  Ctrl+M hotkey registered")
    else:
        log.warning("  Failed to register Ctrl+M hotkey (may be in use by another app)")

    log.info("  Clipboard listener active (event-driven, no polling)")

    msg = ctypes.wintypes.MSG()
    while user32.GetMessageW(ctypes.byref(msg), None, 0, 0) > 0:
        user32.TranslateMessage(ctypes.byref(msg))
        user32.DispatchMessageW(ctypes.byref(msg))

    user32.RemoveClipboardFormatListener(hwnd)
    user32.DestroyWindow(hwnd)


# ── Auto-Install (first run as exe) ───────────────────────────────────
def auto_install():
    """When running as exe, copy to AppData and create startup shortcut if not already installed."""
    if not getattr(sys, "frozen", False):
        return  # Only for .exe

    install_dir = Path(os.environ.get("LOCALAPPDATA", "")) / "ClipboardCoach"
    installed_exe = install_dir / "ClipboardCoach.exe"
    current_exe = Path(sys.executable)

    # Already running from install dir
    if current_exe.parent.resolve() == install_dir.resolve():
        return

    install_dir.mkdir(parents=True, exist_ok=True)

    # Copy exe to install dir
    import shutil
    shutil.copy2(str(current_exe), str(installed_exe))
    log.info("Installed to %s", installed_exe)

    # Create startup shortcut for auto-start at login
    startup_dir = Path(os.environ.get("APPDATA", "")) / "Microsoft" / "Windows" / "Start Menu" / "Programs" / "Startup"
    shortcut_path = startup_dir / "ClipFix.lnk"
    if not shortcut_path.exists():
        try:
            import subprocess
            ps_cmd = (
                f'$ws = New-Object -ComObject WScript.Shell; '
                f'$sc = $ws.CreateShortcut("{shortcut_path}"); '
                f'$sc.TargetPath = "{installed_exe}"; '
                f'$sc.Arguments = "--background"; '
                f'$sc.WorkingDirectory = "{install_dir}"; '
                f'$sc.Description = "ClipFix"; '
                f'$sc.Save()'
            )
            subprocess.run(["powershell", "-Command", ps_cmd],
                           capture_output=True, timeout=10)
            log.info("Auto-start shortcut created")
        except Exception as e:
            log.warning("Could not create startup shortcut: %s", e)


# ── System Tray Icon ──────────────────────────────────────────────────
def _create_tray_icon():
    """Create a green 'CF' icon for the system tray."""
    img = Image.new("RGB", (64, 64), (34, 139, 34))
    draw = ImageDraw.Draw(img)
    try:
        font = ImageFont.truetype("segoeui.ttf", 28)
    except OSError:
        font = ImageFont.load_default()
    draw.text((8, 14), "CF", fill="white", font=font)
    return img


tray_icon = None


def _open_log():
    os.startfile(str(LOG_FILE))


def _quit_app(icon):
    log.info("Quit from tray.")
    icon.stop()
    os._exit(0)


def start_tray_icon():
    """Start the system tray icon in a background thread."""
    global tray_icon
    menu = Menu(
        MenuItem("ClipFix is running", lambda: None, enabled=False),
        Menu.SEPARATOR,
        MenuItem("Open log file", lambda: _open_log()),
        MenuItem("Quit", lambda: _quit_app(tray_icon)),
    )
    tray_icon = Icon("ClipFix", _create_tray_icon(), "ClipFix", menu)
    tray_icon.run_detached()


# ── Main ───────────────────────────────────────────────────────────────
def main():
    global provider

    auto_install()

    try:
        provider = load_provider_from_config()
    except RuntimeError:
        # No provider configured -- show setup wizard
        from setup_wizard import run_setup
        if not run_setup():
            print("Setup cancelled. Exiting.")
            sys.exit(1)
        provider = load_provider_from_config()

    start_tray_icon()

    log.info("-" * 60)
    log.info("  CLIPFIX -- Always-on clipboard text fixer")
    log.info("  Provider: %s", provider.display_name)
    log.info("  Log file: %s", LOG_FILE)
    if BACKGROUND_MODE:
        log.info("  Mode: BACKGROUND")
    else:
        log.info("  Mode: INTERACTIVE")
    log.info("-" * 60)
    log.info("  Copy a message -> get coaching -> Ctrl+M to paste rewrite")
    log.info("  Tray icon active -- right-click to quit")
    log.info("")

    top_patterns = get_top_patterns()
    if top_patterns:
        log.info("  Your recurring patterns: %s", ", ".join(top_patterns))

    try:
        create_clipboard_listener(analyze_in_background)
    except KeyboardInterrupt:
        log.info("ClipFix signing off.")

    if tray_icon:
        tray_icon.stop()


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        log.exception("Fatal error: %s", e)
        # Show error in a message box since console may not be visible
        try:
            import tkinter as tk
            from tkinter import messagebox
            root = tk.Tk()
            root.withdraw()
            messagebox.showerror(
                "ClipFix Error",
                f"{e}\n\nCheck log file:\n{LOG_FILE}",
            )
            root.destroy()
        except Exception:
            pass
        sys.exit(1)
