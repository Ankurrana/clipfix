@echo off
echo ============================================================
echo   ClipFix - Uninstaller
echo ============================================================
echo.

set INSTALL_DIR=%LOCALAPPDATA%\ClipFix

:: Stop running instances
taskkill /f /im ClipFix.exe 2>nul

:: Remove startup shortcut
del "%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup\ClipFix.lnk" 2>nul
echo [OK] Auto-start removed

:: Remove Start Menu shortcut
del "%APPDATA%\Microsoft\Windows\Start Menu\Programs\ClipFix.lnk" 2>nul
echo [OK] Start Menu shortcut removed

:: Remove install directory
if exist "%INSTALL_DIR%" rmdir /s /q "%INSTALL_DIR%"
echo [OK] Files removed

echo.
echo ClipFix has been uninstalled.
pause
