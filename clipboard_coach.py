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
from telemetry import Telemetry

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
    APP_DIR = Path(os.environ.get("LOCALAPPDATA", "")) / "ClipFix"
else:
    # Running as Python script
    APP_DIR = Path(__file__).parent

APP_DIR.mkdir(parents=True, exist_ok=True)
HISTORY_FILE = APP_DIR / "coaching-history.json"
LOG_FILE = APP_DIR / "clipfix.log"

BACKGROUND_MODE = "--background" in sys.argv
telemetry = Telemetry(APP_DIR)

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
# Monotonically increasing ID — each new popup checks if it's still the latest.
# If a newer popup has been requested, the older one auto-dismisses.
_popup_generation = {"value": 0, "lock": threading.Lock()}


def _show_popup(title, msg, duration_ms=6000):
    """Show a non-blocking popup in the bottom-right corner that auto-dismisses."""
    import tkinter as tk

    # Claim a new generation ID so any older popup will self-dismiss
    with _popup_generation["lock"]:
        _popup_generation["value"] += 1
        my_gen = _popup_generation["value"]

    popup = tk.Tk()
    popup.overrideredirect(True)  # No title bar
    popup.attributes("-topmost", True)  # Always on top
    popup.attributes("-alpha", 0.92)  # Slight transparency
    popup.configure(bg="#2d2d2d")

    # Title
    tk.Label(
        popup, text=title, font=("Segoe UI", 11, "bold"),
        fg="#4FC3F7", bg="#2d2d2d", anchor="w",
    ).pack(fill="x", padx=12, pady=(10, 2))

    # Message
    tk.Label(
        popup, text=msg, font=("Segoe UI", 10),
        fg="#ffffff", bg="#2d2d2d", anchor="w", justify="left",
        wraplength=350,
    ).pack(fill="x", padx=12, pady=(0, 10))

    popup.update_idletasks()

    # Position bottom-right of screen
    w = max(popup.winfo_reqwidth(), 380)
    h = popup.winfo_reqheight()
    screen_w = popup.winfo_screenwidth()
    screen_h = popup.winfo_screenheight()
    x = screen_w - w - 20
    y = screen_h - h - 80  # Above taskbar
    popup.geometry(f"{w}x{h}+{x}+{y}")

    # Click to dismiss
    for widget in [popup] + list(popup.winfo_children()):
        widget.bind("<Button-1>", lambda e: popup.destroy())

    def _check_stale():
        """Periodically check if a newer popup has been requested."""
        with _popup_generation["lock"]:
            if _popup_generation["value"] != my_gen:
                popup.destroy()
                return
        popup.after(200, _check_stale)

    # Auto-dismiss after duration
    popup.after(duration_ms, popup.destroy)
    # Poll for staleness (self-dismiss if a newer popup appeared)
    popup.after(200, _check_stale)

    popup.mainloop()


def silent_notify(title, line2, line3=None, duration_ms=6000):
    """Show a popup notification in the bottom-right corner."""
    msg = line2
    if line3:
        msg = f"{line2}\n{line3}"

    log.info("  [notify] %s: %s", title, msg.replace("\n", " | "))

    def _send():
        try:
            _show_popup(title, msg, duration_ms)
        except Exception as e:
            log.warning("  [notify] popup failed: %s", e)

    threading.Thread(target=_send, daemon=True).start()


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


def on_paste_hotkey():
    """Global hotkey: Ctrl+Shift+; pastes the rewrite into the active window."""
    if pending_rewrite["current"] and not pending_rewrite["pasted"]:
        pyperclip.copy(pending_rewrite["current"])
        time.sleep(0.05)
        _simulate_paste()
        pending_rewrite["pasted"] = True
        log.info("  [OK] Rewrite pasted via Ctrl+Shift+;!")
        silent_notify("ClipFix", "Rewrite pasted!")
        telemetry.log_rewrite_pasted(pending_rewrite["current"])


# ── Display ─────────────────────────────────────────────────────────────
def display_result(result):
    if result["verdict"] == "good":
        log.info("[OK] Looks good -- send it.")
        silent_notify("ClipFix", "Looks good -- send it!")
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
            f"Rewrite: {rewrite}  --  Ctrl+Shift+; to paste",
        )
    elif rewrite:
        silent_notify(issue or "Communication Coach", nudge, "Ctrl+Shift+; to paste rewrite")
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

            # Record telemetry
            telemetry.log_analysis(
                input_text=text,
                result=result,
                api_duration=api_duration,
                total_duration=t_total,
                cached=(api_duration == 0.0),
            )

            if rewrite:
                pending_rewrite["current"] = rewrite
                pending_rewrite["pasted"] = False
                log.info("  -> Press Ctrl+Shift+; anywhere to paste the rewrite")
        except Exception as e:
            log.error("[!] Error: %s", e)
        finally:
            analyzing_lock.release()

    if analyzing_lock.acquire(blocking=False):
        log.info("Detected message -- analyzing...")
        threading.Thread(target=_run, daemon=True).start()
    else:
        log.info("  [skip] Analysis already in progress, skipping")


# ── Windows Clipboard Listener + Hotkey ────────────────────────────────
WM_CLIPBOARDUPDATE = 0x031D
WM_HOTKEY = 0x0312
WM_DESTROY = 0x0002
MOD_CONTROL = 0x0002
MOD_SHIFT = 0x0004
VK_OEM_1 = 0xBA  # semicolon key (;/:)
HOTKEY_ID_PASTE = 1

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
        elif msg == WM_HOTKEY and wparam == HOTKEY_ID_PASTE:
            on_paste_hotkey()
            return 0
        elif msg == WM_DESTROY:
            user32.PostQuitMessage(0)
            return 0
        return user32.DefWindowProcW(hwnd, msg, wparam, lparam)

    wnd_proc_cb = WNDPROC(wnd_proc)

    wc = WNDCLASSW()
    wc.lpfnWndProc = ctypes.cast(wnd_proc_cb, ctypes.c_void_p)
    wc.hInstance = kernel32.GetModuleHandleW(None)
    wc.lpszClassName = "ClipFixListener"

    class_atom = user32.RegisterClassW(ctypes.byref(wc))
    if not class_atom:
        raise RuntimeError("Failed to register window class")

    hwnd = user32.CreateWindowExW(
        0, wc.lpszClassName, "ClipFix",
        0, 0, 0, 0, 0,
        None, None, wc.hInstance, None,
    )
    if not hwnd:
        raise RuntimeError("Failed to create listener window")

    if not user32.AddClipboardFormatListener(hwnd):
        raise RuntimeError("Failed to add clipboard listener")

    # Register Ctrl+Shift+; global hotkey
    if user32.RegisterHotKey(hwnd, HOTKEY_ID_PASTE, MOD_CONTROL | MOD_SHIFT, VK_OEM_1):
        log.info("  Ctrl+Shift+; hotkey registered")
    else:
        log.warning("  Failed to register Ctrl+Shift+; hotkey (may be in use by another app)")

    log.info("  Clipboard listener active (event-driven, no polling)")

    msg = ctypes.wintypes.MSG()
    while user32.GetMessageW(ctypes.byref(msg), None, 0, 0) > 0:
        user32.TranslateMessage(ctypes.byref(msg))
        user32.DispatchMessageW(ctypes.byref(msg))

    user32.RemoveClipboardFormatListener(hwnd)
    user32.DestroyWindow(hwnd)


# ── Single Instance (mutex) ────────────────────────────────────────────
MUTEX_NAME = "Global\\ClipFix_SingleInstance"
_mutex_handle = None


def ensure_single_instance():
    """Ensure only one instance of ClipFix is running.

    Uses a Windows named mutex. If another instance holds the mutex,
    terminate it gracefully before proceeding.
    """
    global _mutex_handle
    import subprocess

    # Kill any existing ClipFix processes (except ourselves)
    my_pid = os.getpid()
    try:
        # tasklist outputs lines like: ClipFix.exe    1234 Console  1  25,000 K
        result = subprocess.run(
            ["tasklist", "/FI", "IMAGENAME eq ClipFix.exe", "/FO", "CSV", "/NH"],
            capture_output=True, text=True, timeout=5,
        )
        for line in result.stdout.strip().splitlines():
            parts = line.strip('"').split('","')
            if len(parts) >= 2:
                try:
                    pid = int(parts[1])
                    if pid != my_pid:
                        log.info("  Stopping previous instance (PID %d)...", pid)
                        subprocess.run(
                            ["taskkill", "/PID", str(pid), "/F"],
                            capture_output=True, timeout=5,
                        )
                except (ValueError, IndexError):
                    continue
    except Exception as e:
        log.warning("  Could not check for existing instances: %s", e)

    # Also kill by window class name for dev mode (python clipboard_coach.py)
    # where the process name won't be ClipFix.exe
    try:
        existing_hwnd = user32.FindWindowW("ClipFixListener", None)
        if existing_hwnd:
            log.info("  Found existing ClipFix listener window, sending close...")
            WM_CLOSE = 0x0010
            user32.PostMessageW(existing_hwnd, WM_CLOSE, 0, 0)
            time.sleep(0.5)
    except Exception:
        pass

    # Acquire mutex to prevent future duplicates
    _mutex_handle = ctypes.windll.kernel32.CreateMutexW(None, True, MUTEX_NAME)
    last_error = ctypes.windll.kernel32.GetLastError()
    ERROR_ALREADY_EXISTS = 183
    if last_error == ERROR_ALREADY_EXISTS:
        # Another instance grabbed the mutex between our kill and here — wait briefly
        time.sleep(1)
        _mutex_handle = ctypes.windll.kernel32.CreateMutexW(None, True, MUTEX_NAME)


# ── Registry: Add/Remove Programs ─────────────────────────────────────
UNINSTALL_REG_KEY = r"Software\Microsoft\Windows\CurrentVersion\Uninstall\ClipFix"
APP_VERSION = "1.0.0"


def _register_uninstaller(install_dir: Path, installed_exe: Path):
    """Register ClipFix in Windows 'Apps & Features' / Control Panel."""
    import winreg
    try:
        key = winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, UNINSTALL_REG_KEY, 0,
                                 winreg.KEY_WRITE)
        winreg.SetValueEx(key, "DisplayName", 0, winreg.REG_SZ, "ClipFix")
        winreg.SetValueEx(key, "DisplayVersion", 0, winreg.REG_SZ, APP_VERSION)
        winreg.SetValueEx(key, "Publisher", 0, winreg.REG_SZ, "ClipFix")
        winreg.SetValueEx(key, "InstallLocation", 0, winreg.REG_SZ,
                          str(install_dir))
        winreg.SetValueEx(key, "UninstallString", 0, winreg.REG_SZ,
                          f'"{installed_exe}" --uninstall')
        winreg.SetValueEx(key, "QuietUninstallString", 0, winreg.REG_SZ,
                          f'"{installed_exe}" --uninstall --quiet')
        winreg.SetValueEx(key, "NoModify", 0, winreg.REG_DWORD, 1)
        winreg.SetValueEx(key, "NoRepair", 0, winreg.REG_DWORD, 1)
        # Estimated size in KB
        try:
            size_kb = int(installed_exe.stat().st_size / 1024)
            winreg.SetValueEx(key, "EstimatedSize", 0, winreg.REG_DWORD, size_kb)
        except Exception:
            pass
        winreg.CloseKey(key)
        log.info("  Registered in Apps & Features")
    except Exception as e:
        log.warning("  Could not register uninstaller: %s", e)


def _remove_uninstaller_registry():
    """Remove ClipFix from Windows 'Apps & Features'."""
    import winreg
    try:
        winreg.DeleteKey(winreg.HKEY_CURRENT_USER, UNINSTALL_REG_KEY)
    except FileNotFoundError:
        pass
    except Exception as e:
        log.warning("  Could not remove registry entry: %s", e)


def run_uninstall(quiet=False):
    """Full uninstall: stop app, remove files, shortcuts, and registry entry."""
    import shutil
    import subprocess

    install_dir = Path(os.environ.get("LOCALAPPDATA", "")) / "ClipFix"

    # Stop running instances
    try:
        subprocess.run(["taskkill", "/F", "/IM", "ClipFix.exe"],
                       capture_output=True, timeout=5)
    except Exception:
        pass

    # Remove startup shortcut
    startup_dir = Path(os.environ.get("APPDATA", "")) / "Microsoft" / "Windows" / "Start Menu" / "Programs" / "Startup"
    for name in ("ClipFix.lnk",):
        p = startup_dir / name
        if p.exists():
            p.unlink(missing_ok=True)

    # Remove Start Menu shortcut
    start_menu = Path(os.environ.get("APPDATA", "")) / "Microsoft" / "Windows" / "Start Menu" / "Programs"
    for name in ("ClipFix.lnk",):
        p = start_menu / name
        if p.exists():
            p.unlink(missing_ok=True)

    # Remove registry entry
    _remove_uninstaller_registry()

    # Remove install directory
    if install_dir.exists():
        shutil.rmtree(install_dir, ignore_errors=True)

    if not quiet:
        try:
            import tkinter as tk
            from tkinter import messagebox
            root = tk.Tk()
            root.withdraw()
            messagebox.showinfo("ClipFix", "ClipFix has been uninstalled.")
            root.destroy()
        except Exception:
            pass

    sys.exit(0)


# ── Auto-Install (first run as exe) ───────────────────────────────────
def auto_install():
    """When running as exe, stop any previous instance, clean-install to AppData."""
    if not getattr(sys, "frozen", False):
        return  # Only for .exe

    import shutil
    import subprocess

    install_dir = Path(os.environ.get("LOCALAPPDATA", "")) / "ClipFix"
    installed_exe = install_dir / "ClipFix.exe"
    current_exe = Path(sys.executable)

    # Already running from install dir
    if current_exe.parent.resolve() == install_dir.resolve():
        return

    # Remove previous installation completely
    if install_dir.exists():
        log.info("  Removing previous installation at %s", install_dir)
        # Keep user data files (config, telemetry, history)
        keep = {"config.json", "telemetry.jsonl", "coaching-history.json"}
        for item in install_dir.iterdir():
            if item.name not in keep:
                try:
                    if item.is_dir():
                        shutil.rmtree(item)
                    else:
                        item.unlink()
                except Exception as e:
                    log.warning("  Could not remove %s: %s", item.name, e)

    install_dir.mkdir(parents=True, exist_ok=True)

    # Copy fresh exe
    shutil.copy2(str(current_exe), str(installed_exe))
    log.info("Installed to %s", installed_exe)

    # Create/overwrite startup shortcut for auto-start at login
    startup_dir = Path(os.environ.get("APPDATA", "")) / "Microsoft" / "Windows" / "Start Menu" / "Programs" / "Startup"
    shortcut_path = startup_dir / "ClipFix.lnk"
    try:
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

    # Register in Apps & Features so users can uninstall from Control Panel
    _register_uninstaller(install_dir, installed_exe)


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


def _show_progress():
    """Show a rich 'My Progress' window with key metrics."""
    import tkinter as tk

    all_stats = telemetry.summary()
    week_stats = telemetry.weekly_stats()

    win = tk.Tk()
    win.title("ClipFix - My Progress")
    win.configure(bg="#1e1e1e")
    win.resizable(False, False)
    win.attributes("-topmost", True)

    BG = "#1e1e1e"
    FG = "#e0e0e0"
    CYAN = "#4FC3F7"
    GREEN = "#66BB6A"
    YELLOW = "#FFD54F"
    DIM = "#888888"

    def heading(parent, text):
        tk.Label(parent, text=text, font=("Segoe UI", 13, "bold"),
                 fg=CYAN, bg=BG, anchor="w").pack(fill="x", padx=16, pady=(12, 2))

    def metric(parent, label, value, color=FG):
        f = tk.Frame(parent, bg=BG)
        f.pack(fill="x", padx=24, pady=1)
        tk.Label(f, text=label, font=("Segoe UI", 10), fg=DIM, bg=BG,
                 anchor="w", width=22).pack(side="left")
        tk.Label(f, text=str(value), font=("Segoe UI", 10, "bold"), fg=color,
                 bg=BG, anchor="e").pack(side="right")

    def separator(parent):
        tk.Frame(parent, bg="#333333", height=1).pack(fill="x", padx=16, pady=6)

    # ── This week ──
    heading(win, "This Week (7 days)")
    w_total = week_stats["verdicts"]["good"] + week_stats["verdicts"]["improve"]
    metric(win, "Messages analyzed", w_total)
    metric(win, "Clean (no fixes needed)", f"{week_stats['clean_rate']}%",
           GREEN if week_stats["clean_rate"] >= 70 else YELLOW)
    metric(win, "Rewrites pasted", week_stats["total_rewrites_pasted"])

    # Trend vs previous week
    prev = telemetry.prev_weekly_stats()
    if prev["total_analyses"] > 0:
        delta = week_stats["clean_rate"] - prev["clean_rate"]
        if delta > 0:
            metric(win, "vs. last week", f"+{delta}pp", GREEN)
        elif delta < 0:
            metric(win, "vs. last week", f"{delta}pp", YELLOW)
        else:
            metric(win, "vs. last week", "steady", DIM)

    separator(win)

    # ── All time ──
    heading(win, "All Time")
    metric(win, "Sessions", all_stats["total_sessions"])
    metric(win, "Total analyses", all_stats["total_analyses"])
    metric(win, "Clean rate", f"{all_stats['clean_rate']}%",
           GREEN if all_stats["clean_rate"] >= 70 else YELLOW)
    if all_stats["verdicts"]["improve"] > 0:
        metric(win, "Acceptance rate", f"{all_stats['acceptance_rate']}%")

    separator(win)

    # ── Top issues ──
    if all_stats["top_issues"]:
        heading(win, "Top Issues")
        max_count = all_stats["top_issues"][0][1] if all_stats["top_issues"] else 1
        for issue, count in all_stats["top_issues"][:5]:
            f = tk.Frame(win, bg=BG)
            f.pack(fill="x", padx=24, pady=1)
            tk.Label(f, text=issue, font=("Segoe UI", 10), fg=FG, bg=BG,
                     anchor="w", width=20).pack(side="left")
            # Text bar
            bar_len = max(1, int(count / max_count * 12))
            bar = "\u2588" * bar_len
            tk.Label(f, text=f"{bar} {count}", font=("Consolas", 9),
                     fg=CYAN, bg=BG, anchor="w").pack(side="left", padx=(4, 0))

    # ── Close button ──
    tk.Button(win, text="Close", command=win.destroy,
              bg="#333333", fg=FG, font=("Segoe UI", 9),
              relief="flat", padx=16, pady=4,
              ).pack(pady=(10, 12))

    # Center on screen
    win.update_idletasks()
    w = max(win.winfo_reqwidth(), 360)
    h = win.winfo_reqheight()
    x = (win.winfo_screenwidth() - w) // 2
    y = (win.winfo_screenheight() - h) // 2
    win.geometry(f"{w}x{h}+{x}+{y}")

    win.mainloop()


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
        MenuItem("My Progress", lambda: threading.Thread(
            target=_show_progress, daemon=True).start()),
        MenuItem("Open log file", lambda: _open_log()),
        MenuItem("Quit", lambda: _quit_app(tray_icon)),
    )
    tray_icon = Icon("ClipFix", _create_tray_icon(), "ClipFix", menu)
    tray_icon.run_detached()
    # Wait for tray icon to initialize
    for _ in range(20):
        if tray_icon.visible:
            break
        time.sleep(0.25)
    log.info("  Tray icon visible: %s", tray_icon.visible)


# ── Main ───────────────────────────────────────────────────────────────
def main():
    global provider

    ensure_single_instance()
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

    telemetry.log_session_start(provider.display_name)

    log.info("-" * 60)
    log.info("  CLIPFIX -- Always-on clipboard text fixer")
    log.info("  Provider: %s", provider.display_name)
    log.info("  Log file: %s", LOG_FILE)
    if BACKGROUND_MODE:
        log.info("  Mode: BACKGROUND")
    else:
        log.info("  Mode: INTERACTIVE")
    log.info("-" * 60)
    log.info("  Copy a message -> get coaching -> Ctrl+Shift+; to paste rewrite")
    log.info("  Tray icon active -- right-click to quit")
    log.info("")

    top_patterns = get_top_patterns()
    if top_patterns:
        log.info("  Your recurring patterns: %s", ", ".join(top_patterns))

    # Startup summary: show weekly digest if due, else a quick one-liner
    if telemetry.should_show_weekly_digest():
        digest = telemetry.weekly_digest()
        if digest:
            silent_notify("ClipFix - Weekly Digest", digest, duration_ms=12000)
            telemetry.mark_weekly_digest_shown()
            log.info("  [digest] Weekly digest shown")
        else:
            silent_notify("ClipFix", "Running! Copy a message to get started.")
    else:
        startup_msg = telemetry.startup_summary()
        if startup_msg:
            silent_notify("ClipFix", startup_msg, duration_ms=8000)
        else:
            silent_notify("ClipFix", "Running! Copy a message to get started.")

    try:
        create_clipboard_listener(analyze_in_background)
    except KeyboardInterrupt:
        log.info("ClipFix signing off.")

    if tray_icon:
        tray_icon.stop()


if __name__ == "__main__":
    if "--uninstall" in sys.argv:
        run_uninstall(quiet="--quiet" in sys.argv)

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
