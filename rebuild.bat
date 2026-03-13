@echo off
echo ============================================================
echo   ClipFix - Rebuild and Reinstall
echo ============================================================
echo.

:: Stop running instances
taskkill /f /im ClipFix.exe 2>nul
timeout /t 2 /nobreak >nul

:: Rebuild
echo [1/3] Building executable...
python build.py
if %errorlevel% neq 0 (
    echo BUILD FAILED!
    pause
    exit /b 1
)

:: Run tests
echo.
echo [2/3] Running tests...
python test_integration.py
if %errorlevel% neq 0 (
    echo TESTS FAILED! Install anyway? (y/n)
    set /p CONT=
    if /i not "%CONT%"=="y" exit /b 1
)

:: Reinstall
echo.
echo [3/3] Installing...
set INSTALL_DIR=%LOCALAPPDATA%\ClipFix
if not exist "%INSTALL_DIR%" mkdir "%INSTALL_DIR%"
copy /y dist\ClipFix.exe "%INSTALL_DIR%\ClipFix.exe" >nul
if exist config.json copy /y config.json "%INSTALL_DIR%\config.json" >nul

echo.
echo ============================================================
echo   Done! Updated executable installed.
echo ============================================================
echo.
set /p LAUNCH="Launch now? (y/n): "
if /i "%LAUNCH%"=="y" start "" "%INSTALL_DIR%\ClipFix.exe"
pause
