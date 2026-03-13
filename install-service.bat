@echo off
echo ============================================================
echo   ClipFix -- Service Installer
echo ============================================================
echo.

:: Get Python path
for /f "delims=" %%i in ('where pythonw 2^>nul') do set PYTHONW=%%i
if "%PYTHONW%"=="" (
    echo ERROR: pythonw.exe not found. Make sure Python is installed and on PATH.
    pause
    exit /b 1
)

:: Check API key
if "%AZURE_OPENAI_API_KEY%"=="" (
    echo ERROR: AZURE_OPENAI_API_KEY environment variable not set.
    echo   Run: setx AZURE_OPENAI_API_KEY "your-key-here"
    pause
    exit /b 1
)

set SCRIPT_DIR=%~dp0
set SCRIPT=%SCRIPT_DIR%clipboard_coach.py

echo Python:  %PYTHONW%
echo Script:  %SCRIPT%
echo Mode:    Background (auto-copy rewrites)
echo.

:: Delete existing task if present
schtasks /delete /tn "ClipFix" /f >nul 2>&1

:: Create scheduled task to run at logon
schtasks /create ^
    /tn "ClipFix" ^
    /tr "\"%PYTHONW%\" \"%SCRIPT%\" --background" ^
    /sc onlogon ^
    /rl highest ^
    /f

if %errorlevel% equ 0 (
    echo.
    echo [OK] Scheduled task "ClipFix" created successfully.
    echo     It will start automatically when you log in.
    echo.
    echo To start it now:
    echo     schtasks /run /tn "ClipFix"
    echo.
    echo To stop it:
    echo     taskkill /f /im pythonw.exe
    echo.
    echo To remove it:
    echo     schtasks /delete /tn "ClipFix" /f
    echo.
    echo Log file: %SCRIPT_DIR%clipboard-coach.log
) else (
    echo.
    echo [!] Failed to create task. Try running this script as Administrator.
)

pause
