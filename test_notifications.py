"""Notification diagnostic script for ClipFix.

Copy this single file to the customer's machine and run:
    python test_notifications.py

Tests every notification method available on Windows to find what works.
No dependencies required beyond Python stdlib + pywin32 (optional).
"""
import os
import sys
import time
import ctypes
import ctypes.wintypes
import subprocess

print("=" * 60)
print("  ClipFix Notification Diagnostics")
print(f"  Python: {sys.version.split()[0]}")
print(f"  OS: {sys.platform}")
print("=" * 60)
print()

results = {}


def test(name):
    print(f"  [{name}] ", end="", flush=True)
    return name


# ── 1. Windows version ────────────────────────────────────────────────
name = test("Windows version")
try:
    ver = sys.getwindowsversion()
    print(f"Windows {ver.major}.{ver.minor}.{ver.build}")
    results[name] = "OK"
except Exception as e:
    print(f"FAILED: {e}")
    results[name] = "FAIL"

# ── 2. Notification settings (registry) ───────────────────────────────
name = test("Notifications enabled (registry)")
try:
    import winreg
    key = winreg.OpenKey(
        winreg.HKEY_CURRENT_USER,
        r"Software\Microsoft\Windows\CurrentVersion\PushNotifications",
    )
    try:
        val, _ = winreg.QueryValueEx(key, "ToastEnabled")
        if val == 1:
            print("YES (ToastEnabled=1)")
            results[name] = "OK"
        else:
            print(f"NO (ToastEnabled={val}) -- notifications are DISABLED")
            results[name] = "BLOCKED"
    except FileNotFoundError:
        print("YES (no override, default enabled)")
        results[name] = "OK"
    finally:
        winreg.CloseKey(key)
except Exception as e:
    print(f"Could not check: {e}")
    results[name] = "UNKNOWN"

# ── 3. Focus Assist / Do Not Disturb ─────────────────────────────────
name = test("Focus Assist / DND")
try:
    result = subprocess.run(
        ["powershell", "-Command",
         "(Get-ItemProperty -Path 'HKCU:\\Software\\Microsoft\\Windows\\CurrentVersion\\CloudStore\\Store\\DefaultAccount\\Current\\default$windows.data.notifications.quiethourssettings\\windows.data.notifications.quiethourssettings' -ErrorAction SilentlyContinue).Data"],
        capture_output=True, text=True, timeout=10,
    )
    if result.stdout.strip():
        print("ACTIVE -- may suppress notifications")
        results[name] = "WARN"
    else:
        print("Off")
        results[name] = "OK"
except Exception as e:
    print(f"Could not check: {e}")
    results[name] = "UNKNOWN"

# ── 4. Group Policy restrictions ──────────────────────────────────────
name = test("Group Policy toast restriction")
try:
    import winreg
    blocked = False
    for path in [
        r"SOFTWARE\Policies\Microsoft\Windows\Explorer",
        r"SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer",
    ]:
        try:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, path)
            val, _ = winreg.QueryValueEx(key, "DisableNotificationCenter")
            winreg.CloseKey(key)
            if val == 1:
                print(f"BLOCKED by policy ({path})")
                blocked = True
                results[name] = "BLOCKED"
                break
        except (FileNotFoundError, OSError):
            pass
    # Also check machine policy
    if not blocked:
        try:
            key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE,
                                 r"SOFTWARE\Policies\Microsoft\Windows\Explorer")
            val, _ = winreg.QueryValueEx(key, "DisableNotificationCenter")
            winreg.CloseKey(key)
            if val == 1:
                print("BLOCKED by machine policy")
                blocked = True
                results[name] = "BLOCKED"
        except (FileNotFoundError, OSError):
            pass
    if not blocked:
        print("No restrictions found")
        results[name] = "OK"
except Exception as e:
    print(f"Could not check: {e}")
    results[name] = "UNKNOWN"

# ── 5. Test: MessageBox (always works) ────────────────────────────────
name = test("MessageBox (Win32)")
try:
    MB_OK = 0x0
    MB_ICONINFORMATION = 0x40
    MB_SYSTEMMODAL = 0x1000
    ctypes.windll.user32.MessageBoxW(
        0, "If you see this popup, basic Win32 UI works.\n\nClick OK to continue.",
        "ClipFix Diagnostic", MB_OK | MB_ICONINFORMATION,
    )
    print("OK (shown)")
    results[name] = "OK"
except Exception as e:
    print(f"FAILED: {e}")
    results[name] = "FAIL"

# ── 6. Test: Balloon notification via Shell_NotifyIconW ───────────────
name = test("Balloon notification (Shell_NotifyIconW)")
try:
    class NOTIFYICONDATAW(ctypes.Structure):
        _fields_ = [
            ("cbSize", ctypes.wintypes.DWORD),
            ("hWnd", ctypes.wintypes.HWND),
            ("uID", ctypes.wintypes.UINT),
            ("uFlags", ctypes.wintypes.UINT),
            ("uCallbackMessage", ctypes.wintypes.UINT),
            ("hIcon", ctypes.wintypes.HICON),
            ("szTip", ctypes.c_wchar * 128),
            ("dwState", ctypes.wintypes.DWORD),
            ("dwStateMask", ctypes.wintypes.DWORD),
            ("szInfo", ctypes.c_wchar * 256),
            ("uVersion", ctypes.wintypes.UINT),
            ("szInfoTitle", ctypes.c_wchar * 64),
            ("dwInfoFlags", ctypes.wintypes.DWORD),
        ]

    NIF_ICON = 0x02
    NIF_TIP = 0x04
    NIF_INFO = 0x10
    NIM_ADD = 0x00
    NIM_MODIFY = 0x01
    NIM_DELETE = 0x02
    NIIF_INFO = 0x01
    NIIF_NOSOUND = 0x10

    shell32 = ctypes.windll.shell32
    user32 = ctypes.windll.user32

    hIcon = user32.LoadIconW(None, ctypes.cast(32516, ctypes.wintypes.LPCWSTR))  # IDI_INFORMATION

    nid = NOTIFYICONDATAW()
    nid.cbSize = ctypes.sizeof(NOTIFYICONDATAW)
    nid.hWnd = 0
    nid.uID = 9999
    nid.uFlags = NIF_ICON | NIF_TIP | NIF_INFO
    nid.hIcon = hIcon
    nid.szTip = "ClipFix Diagnostic"
    nid.szInfoTitle = "ClipFix Test"
    nid.szInfo = "If you see this balloon, Shell_NotifyIcon works!"
    nid.dwInfoFlags = NIIF_INFO | NIIF_NOSOUND

    added = shell32.Shell_NotifyIconW(NIM_ADD, ctypes.byref(nid))
    if added:
        print("OK (check taskbar for balloon)")
        results[name] = "OK"
        time.sleep(4)
        shell32.Shell_NotifyIconW(NIM_DELETE, ctypes.byref(nid))
    else:
        print("FAILED (Shell_NotifyIconW returned 0)")
        results[name] = "FAIL"
except Exception as e:
    print(f"FAILED: {e}")
    results[name] = "FAIL"

# ── 7. Test: pystray balloon ──────────────────────────────────────────
name = test("pystray tray balloon")
try:
    from PIL import Image, ImageDraw
    from pystray import Icon, Menu, MenuItem

    img = Image.new("RGB", (64, 64), (34, 139, 34))
    draw = ImageDraw.Draw(img)
    draw.text((8, 14), "CF", fill="white")

    icon = Icon("ClipFixDiag", img, "ClipFix Diagnostic",
                Menu(MenuItem("Test", lambda: None)))
    icon.run_detached()
    time.sleep(1.5)

    if icon.visible:
        icon.notify("If you see this, pystray notifications work!", "ClipFix Test")
        print("OK (check taskbar for balloon)")
        results[name] = "OK"
        time.sleep(4)
    else:
        print("FAILED (tray icon not visible)")
        results[name] = "FAIL"
    icon.stop()
except ImportError:
    print("SKIPPED (pystray/pillow not installed)")
    results[name] = "SKIP"
except Exception as e:
    print(f"FAILED: {e}")
    results[name] = "FAIL"

# ── 8. Test: WinRT toast notification ─────────────────────────────────
name = test("WinRT toast notification")
try:
    result = subprocess.run(
        ["powershell", "-Command", """
[Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime] | Out-Null
$template = [Windows.UI.Notifications.ToastNotificationManager]::GetTemplateContent([Windows.UI.Notifications.ToastTemplateType]::ToastText02)
$textNodes = $template.GetElementsByTagName("text")
$textNodes.Item(0).AppendChild($template.CreateTextNode("ClipFix Test")) | Out-Null
$textNodes.Item(1).AppendChild($template.CreateTextNode("If you see this toast, WinRT notifications work!")) | Out-Null
$toast = [Windows.UI.Notifications.ToastNotification]::new($template)
[Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier("ClipFix").Show($toast)
Write-Output "OK"
"""],
        capture_output=True, text=True, timeout=15,
    )
    if "OK" in result.stdout:
        print("OK (check for toast popup)")
        results[name] = "OK"
        time.sleep(3)
    else:
        err = result.stderr.strip().split("\n")[0] if result.stderr else "Unknown error"
        print(f"FAILED: {err}")
        results[name] = "FAIL"
except Exception as e:
    print(f"FAILED: {e}")
    results[name] = "FAIL"

# ── Summary ───────────────────────────────────────────────────────────
print()
print("=" * 60)
print("  Summary")
print("=" * 60)

blocked = [k for k, v in results.items() if v == "BLOCKED"]
failed = [k for k, v in results.items() if v == "FAIL"]
ok = [k for k, v in results.items() if v == "OK"]

if blocked:
    print()
    print("  BLOCKED BY POLICY:")
    for b in blocked:
        print(f"    - {b}")
    print()
    print("  -> Contact IT to enable notifications, or use ClipFix")
    print("     in console mode (python clipboard_coach.py) to see")
    print("     results in the terminal instead of notifications.")

if failed:
    print()
    print("  FAILED:")
    for f in failed:
        print(f"    - {f}")

# Recommend the best working method
notification_tests = [
    "WinRT toast notification",
    "pystray tray balloon",
    "Balloon notification (Shell_NotifyIconW)",
]
working = [t for t in notification_tests if results.get(t) == "OK"]

print()
if working:
    print(f"  WORKING NOTIFICATION METHODS: {len(working)}")
    for w in working:
        print(f"    + {w}")
    print()
    print(f"  -> Best method for ClipFix: {working[0]}")
else:
    print("  NO NOTIFICATION METHOD WORKS on this machine.")
    print("  -> Notifications may be disabled by Group Policy.")
    print("  -> Run ClipFix in console mode to see results in terminal.")

print()
print("=" * 60)
