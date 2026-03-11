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

import keyboard
import pyperclip
import win32clipboard
from win11toast import toast

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
SCRIPT_DIR = Path(__file__).parent
HISTORY_FILE = SCRIPT_DIR / "coaching-history.json"
LOG_FILE = SCRIPT_DIR / "clipboard-coach.log"

BACKGROUND_MODE = "--background" in sys.argv

# ── Logging ─────────────────────────────────────────────────────────────
if BACKGROUND_MODE:
    logging.basicConfig(
        filename=str(LOG_FILE),
        level=logging.INFO,
        format="%(asctime)s %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )
else:
    logging.basicConfig(
        level=logging.INFO,
        format="%(message)s",
    )
log = logging.getLogger("coach")


# ── Silent Notification (non-blocking) ─────────────────────────────────
def _escape_xml(text):
    return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")


def silent_notify(title, line2, line3=None, scenario="reminder"):
    def _send():
        try:
            t = _escape_xml(title)
            l2 = _escape_xml(line2)
            l3_xml = f"<text>{_escape_xml(line3)}</text>" if line3 else ""
            xml = f'''<toast activationType="protocol" launch="http:" scenario="{scenario}">
    <visual>
        <binding template="ToastGeneric">
            <text>{t}</text>
            <text>{l2}</text>
            {l3_xml}
        </binding>
    </visual>
    <audio silent="true" />
</toast>'''
            toast(xml=xml)
        except Exception:
            pass
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


def on_ctrl_m():
    """Global hotkey: Ctrl+M pastes the rewrite into the active window."""
    if pending_rewrite["current"] and not pending_rewrite["pasted"]:
        pyperclip.copy(pending_rewrite["current"])
        time.sleep(0.05)
        keyboard.send("ctrl+v")
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


# ── Windows Clipboard Listener ─────────────────────────────────────────
WM_CLIPBOARDUPDATE = 0x031D
WM_DESTROY = 0x0002

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

    log.info("  Clipboard listener active (event-driven, no polling)")

    msg = ctypes.wintypes.MSG()
    while user32.GetMessageW(ctypes.byref(msg), None, 0, 0) > 0:
        user32.TranslateMessage(ctypes.byref(msg))
        user32.DispatchMessageW(ctypes.byref(msg))

    user32.RemoveClipboardFormatListener(hwnd)
    user32.DestroyWindow(hwnd)


# ── Main ───────────────────────────────────────────────────────────────
def main():
    global provider

    try:
        provider = load_provider_from_config()
    except RuntimeError:
        # No provider configured -- show setup wizard
        from setup_wizard import run_setup
        if not run_setup():
            print("Setup cancelled. Exiting.")
            sys.exit(1)
        provider = load_provider_from_config()

    keyboard.add_hotkey("ctrl+m", on_ctrl_m, suppress=True)

    log.info("-" * 60)
    log.info("  CLIPBOARD COACH -- Always-on communication improvement")
    log.info("  Provider: %s", provider.display_name)
    if BACKGROUND_MODE:
        log.info("  Mode: BACKGROUND")
        log.info("  Log file: %s", LOG_FILE)
    else:
        log.info("  Mode: INTERACTIVE")
    log.info("-" * 60)
    log.info("  Copy a message -> get coaching -> Ctrl+M to paste rewrite")
    log.info("  Press Ctrl+C to quit.")
    log.info("")

    top_patterns = get_top_patterns()
    if top_patterns:
        log.info("  Your recurring patterns: %s", ", ".join(top_patterns))

    try:
        create_clipboard_listener(analyze_in_background)
    except KeyboardInterrupt:
        log.info("Coach signing off. Keep communicating with impact!")

    keyboard.unhook_all()


if __name__ == "__main__":
    main()
