@echo off
echo ============================================================
echo   Clipboard Coach - Uninstaller
echo ============================================================
echo.

set INSTALL_DIR=%LOCALAPPDATA%\ClipboardCoach

:: Stop running instances
taskkill /f /im ClipboardCoach.exe 2>nul

:: Remove startup shortcut
del "%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup\Clipboard Coach.lnk" 2>nul
echo [OK] Auto-start removed

:: Remove Start Menu shortcut
del "%APPDATA%\Microsoft\Windows\Start Menu\Programs\Clipboard Coach.lnk" 2>nul
echo [OK] Start Menu shortcut removed

:: Remove install directory
if exist "%INSTALL_DIR%" rmdir /s /q "%INSTALL_DIR%"
echo [OK] Files removed

echo.
echo Clipboard Coach has been uninstalled.
pause
