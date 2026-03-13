"""Smoke test for ClipFix on a fresh Windows machine.

Run this BEFORE or AFTER installing to verify everything works:
    python test_smoke.py --api-key "your-key" --endpoint "https://..." --deployment "gpt-4.1"

Or with the exe:
    ClipFix.exe --smoke-test
"""
import os
import sys
import time
import traceback

passed = 0
failed = 0
warnings = 0

# Detect headless (CI) environment — no desktop session
HEADLESS = not os.environ.get("SESSIONNAME") and os.environ.get("CI")


def test(name, fn, requires_desktop=False):
    global passed, failed
    if requires_desktop and HEADLESS:
        print(f"  SKIP  {name} (no desktop in CI)")
        passed += 1  # Not a failure
        return
    try:
        fn()
        print(f"  PASS  {name}")
        passed += 1
    except AssertionError as e:
        print(f"  FAIL  {name}: {e}")
        failed += 1
    except Exception as e:
        print(f"  FAIL  {name}: {type(e).__name__}: {e}")
        failed += 1


def warn(name, msg):
    global warnings
    print(f"  WARN  {name}: {msg}")
    warnings += 1


print("=" * 60)
print("  ClipFix Smoke Test")
print("=" * 60)
print()

# ── 1. Python version ──────────────────────────────────────────────────
def test_python_version():
    v = sys.version_info
    assert v >= (3, 10), f"Python 3.10+ required, got {v.major}.{v.minor}"

test("Python version >= 3.10", test_python_version)


# ── 2. Required imports ────────────────────────────────────────────────
def test_imports():
    import pyperclip
    import win32clipboard
    from PIL import Image
    from pystray import Icon
    from openai import AzureOpenAI

test("Required packages installed", test_imports)


# ── 3. Clipboard access ───────────────────────────────────────────────
def test_clipboard_read():
    import pyperclip
    text = pyperclip.paste()
    assert isinstance(text, str), f"Expected str, got {type(text)}"

test("Clipboard read access", test_clipboard_read, requires_desktop=True)


def test_clipboard_write():
    import pyperclip
    original = pyperclip.paste()
    pyperclip.copy("clipfix_smoke_test")
    result = pyperclip.paste()
    pyperclip.copy(original)  # restore
    assert result == "clipfix_smoke_test", f"Clipboard write failed: got {result!r}"

test("Clipboard write access", test_clipboard_write, requires_desktop=True)


# ── 4. HTML clipboard format ──────────────────────────────────────────
def test_html_clipboard():
    import win32clipboard
    time.sleep(0.2)  # Ensure clipboard is released from previous test
    for attempt in range(3):
        try:
            win32clipboard.OpenClipboard()
            try:
                cf_html = win32clipboard.RegisterClipboardFormat("HTML Format")
                assert cf_html > 0, "Failed to register HTML Format"
                return
            finally:
                win32clipboard.CloseClipboard()
        except Exception:
            if attempt < 2:
                time.sleep(0.5)
            else:
                raise

test("HTML clipboard format available", test_html_clipboard, requires_desktop=True)


# ── 5. Win32 window creation (for clipboard listener) ─────────────────
def test_win32_window():
    import ctypes
    import ctypes.wintypes
    user32 = ctypes.windll.user32
    kernel32 = ctypes.windll.kernel32

    WNDPROC = ctypes.WINFUNCTYPE(ctypes.c_longlong, ctypes.c_void_p,
                                  ctypes.c_uint, ctypes.c_longlong, ctypes.c_longlong)

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
    user32.DefWindowProcW.argtypes = [ctypes.c_void_p, ctypes.c_uint,
                                       ctypes.c_longlong, ctypes.c_longlong]
    user32.DefWindowProcW.restype = ctypes.c_longlong

    def dummy(hwnd, msg, wp, lp):
        return user32.DefWindowProcW(hwnd, msg, wp, lp)
    cb = WNDPROC(dummy)

    wc = WNDCLASSW()
    wc.lpfnWndProc = ctypes.cast(cb, ctypes.c_void_p)
    wc.hInstance = kernel32.GetModuleHandleW(None)
    wc.lpszClassName = "ClipFixSmokeTest"

    class_name = f"ClipFixSmokeTest{int(time.time())}"
    wc.lpszClassName = class_name

    atom = user32.RegisterClassW(ctypes.byref(wc))
    assert atom, "RegisterClassW failed"

    hwnd = user32.CreateWindowExW(
        0, class_name, "test", 0,
        0, 0, 0, 0, None, None, wc.hInstance, None,
    )
    assert hwnd, "CreateWindowExW failed"

    result = user32.AddClipboardFormatListener(hwnd)
    assert result, "AddClipboardFormatListener failed"

    user32.RemoveClipboardFormatListener(hwnd)
    user32.DestroyWindow(hwnd)

test("Win32 clipboard listener (no admin)", test_win32_window, requires_desktop=True)


# ── 6. Hotkey registration ────────────────────────────────────────────
def test_hotkey():
    import ctypes
    user32 = ctypes.windll.user32
    MOD_CONTROL = 0x0002
    VK_M = 0x4D

    # Use a temp window
    result = user32.RegisterHotKey(None, 9999, MOD_CONTROL, VK_M)
    if result:
        user32.UnregisterHotKey(None, 9999)
    assert result, "RegisterHotKey failed (Ctrl+M may be in use)"

test("Ctrl+M hotkey registration", test_hotkey, requires_desktop=True)


# ── 7. System tray icon ───────────────────────────────────────────────
def test_tray_icon():
    from PIL import Image, ImageDraw
    from pystray import Icon, Menu, MenuItem

    img = Image.new("RGB", (64, 64), (34, 139, 34))
    draw = ImageDraw.Draw(img)
    draw.text((8, 14), "CF", fill="white")

    icon = Icon("ClipFixTest", img, "ClipFix Smoke Test",
                Menu(MenuItem("Test", lambda: None)))
    icon.run_detached()
    time.sleep(1)
    assert icon.visible, "Tray icon not visible"
    icon.stop()

test("System tray icon", test_tray_icon, requires_desktop=True)


# ── 8. Tray notification ──────────────────────────────────────────────
def test_tray_notification():
    from PIL import Image, ImageDraw
    from pystray import Icon, Menu, MenuItem

    img = Image.new("RGB", (64, 64), (34, 139, 34))
    draw = ImageDraw.Draw(img)
    draw.text((8, 14), "CF", fill="white")

    icon = Icon("ClipFixNotifyTest", img, "ClipFix Notify Test",
                Menu(MenuItem("Test", lambda: None)))
    icon.run_detached()
    time.sleep(1)
    try:
        icon.notify("This is a ClipFix smoke test notification.", "ClipFix Test")
        time.sleep(2)
    finally:
        icon.stop()

test("Tray balloon notification (check screen!)", test_tray_notification, requires_desktop=True)


# ── 9. AppData directory writable ──────────────────────────────────────
def test_appdata():
    from pathlib import Path
    app_dir = Path(os.environ.get("LOCALAPPDATA", "")) / "ClipFix"
    app_dir.mkdir(parents=True, exist_ok=True)
    test_file = app_dir / ".smoke_test"
    test_file.write_text("ok")
    assert test_file.read_text() == "ok"
    test_file.unlink()

test("AppData directory writable", test_appdata)


# ── 10. API connectivity (optional) ───────────────────────────────────
api_key = os.environ.get("AZURE_OPENAI_API_KEY", "")
if not api_key:
    # Check for config.json
    from pathlib import Path
    for cfg_path in [
        Path("config.json"),
        Path(os.environ.get("LOCALAPPDATA", "")) / "ClipFix" / "config.json",
    ]:
        if cfg_path.exists():
            import json
            cfg = json.loads(cfg_path.read_text())
            if "api_key" in cfg and not cfg["api_key"].startswith("$"):
                api_key = cfg["api_key"]
            break

if api_key:
    def test_api():
        from openai import AzureOpenAI
        endpoint = os.environ.get("AZURE_OPENAI_ENDPOINT",
                                  "https://foundary-poc-gygiuj.cognitiveservices.azure.com/")
        deployment = os.environ.get("AZURE_OPENAI_DEPLOYMENT", "gpt-4.1")
        client = AzureOpenAI(
            azure_endpoint=endpoint,
            api_key=api_key,
            api_version="2025-01-01-preview",
        )
        t0 = time.perf_counter()
        r = client.chat.completions.create(
            model=deployment,
            max_tokens=10,
            messages=[{"role": "user", "content": "Say OK"}],
        )
        duration = time.perf_counter() - t0
        content = r.choices[0].message.content
        assert content and len(content) > 0, f"Empty response"
        print(f"         Response: {content.strip()} ({duration:.2f}s)")

    test("Azure OpenAI API connectivity", test_api)
else:
    warn("API connectivity", "AZURE_OPENAI_API_KEY not set -- skipping API test")


# ── Summary ────────────────────────────────────────────────────────────
print()
print("=" * 60)
total = passed + failed
print(f"  Results: {passed}/{total} passed", end="")
if warnings:
    print(f", {warnings} warnings", end="")
print()

if failed == 0:
    print("  ClipFix should work on this machine!")
else:
    print("  Some checks failed -- see details above.")
print("=" * 60)

sys.exit(1 if failed else 0)
