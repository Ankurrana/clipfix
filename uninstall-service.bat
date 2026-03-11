@echo off
echo Stopping Clipboard Coach...
taskkill /f /im pythonw.exe 2>nul
echo Removing scheduled task...
schtasks /delete /tn "ClipboardCoach" /f 2>nul
echo [OK] Clipboard Coach service removed.
pause
