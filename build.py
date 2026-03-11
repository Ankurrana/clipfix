"""Build Clipboard Coach into a Windows executable."""
import subprocess
import sys
from pathlib import Path

SCRIPT_DIR = Path(__file__).parent

def build():
    print("Building Clipboard Coach...")
    print("=" * 60)

    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--name", "ClipboardCoach",
        "--onefile",
        "--windowed",              # No console window (uses notifications)
        "--uac-admin",             # Request admin elevation via UAC prompt
        "--manifest", str(SCRIPT_DIR / "clipboard_coach.manifest"),
        "--add-data", f"{SCRIPT_DIR / 'providers.py'};.",
        "--add-data", f"{SCRIPT_DIR / 'setup_wizard.py'};.",
        "--add-data", f"{SCRIPT_DIR / 'config.example.json'};.",
        "--hidden-import", "openai",
        "--hidden-import", "anthropic",
        "--hidden-import", "win32clipboard",
        "--hidden-import", "pyperclip",
        "--hidden-import", "tkinter",
        "--hidden-import", "pystray",
        "--hidden-import", "PIL",
        str(SCRIPT_DIR / "clipboard_coach.py"),
    ]

    result = subprocess.run(cmd, cwd=str(SCRIPT_DIR))

    if result.returncode == 0:
        exe_path = SCRIPT_DIR / "dist" / "ClipboardCoach.exe"
        print()
        print("=" * 60)
        print(f"  BUILD SUCCESSFUL!")
        print(f"  Executable: {exe_path}")
        print(f"  Size: {exe_path.stat().st_size / 1024 / 1024:.1f} MB")
        print("=" * 60)
    else:
        print("BUILD FAILED!")
        sys.exit(1)


if __name__ == "__main__":
    build()
