@echo off
echo ============================================================
echo   Clipboard Coach - Installer
echo ============================================================
echo.

set INSTALL_DIR=%LOCALAPPDATA%\ClipboardCoach
set EXE_SRC=%~dp0dist\ClipboardCoach.exe

if not exist "%EXE_SRC%" (
    echo ERROR: ClipboardCoach.exe not found. Run "python build.py" first.
    pause
    exit /b 1
)

echo Installing to: %INSTALL_DIR%
echo.

:: Create install directory
if not exist "%INSTALL_DIR%" mkdir "%INSTALL_DIR%"

:: Copy files
copy /y "%EXE_SRC%" "%INSTALL_DIR%\ClipboardCoach.exe" >nul
if exist "%~dp0config.json" copy /y "%~dp0config.json" "%INSTALL_DIR%\config.json" >nul
copy /y "%~dp0config.example.json" "%INSTALL_DIR%\config.example.json" >nul

:: Create Start Menu shortcut
set SHORTCUT_DIR=%APPDATA%\Microsoft\Windows\Start Menu\Programs
powershell -Command "$ws = New-Object -ComObject WScript.Shell; $sc = $ws.CreateShortcut('%SHORTCUT_DIR%\Clipboard Coach.lnk'); $sc.TargetPath = '%INSTALL_DIR%\ClipboardCoach.exe'; $sc.WorkingDirectory = '%INSTALL_DIR%'; $sc.Description = 'Always-on communication coach'; $sc.Save()"
echo [OK] Start Menu shortcut created

:: Create Startup shortcut (auto-start at login)
set STARTUP_DIR=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup
powershell -Command "$ws = New-Object -ComObject WScript.Shell; $sc = $ws.CreateShortcut('%STARTUP_DIR%\Clipboard Coach.lnk'); $sc.TargetPath = '%INSTALL_DIR%\ClipboardCoach.exe'; $sc.Arguments = '--background'; $sc.WorkingDirectory = '%INSTALL_DIR%'; $sc.Description = 'Clipboard Coach (background)'; $sc.Save()"
echo [OK] Auto-start at login enabled

echo.
echo ============================================================
echo   Installation complete!
echo.
echo   Location:  %INSTALL_DIR%
echo   Start:     Search "Clipboard Coach" in Start Menu
echo   Auto-start: Runs at login (background mode)
echo.
echo   On first run, a setup wizard will ask for your
echo   LLM provider and API key.
echo ============================================================
echo.

:: Ask to launch now
set /p LAUNCH="Launch Clipboard Coach now? (y/n): "
if /i "%LAUNCH%"=="y" start "" "%INSTALL_DIR%\ClipboardCoach.exe"

pause
