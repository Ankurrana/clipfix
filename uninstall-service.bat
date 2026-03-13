@echo off
echo Stopping ClipFix...
taskkill /f /im pythonw.exe 2>nul
echo Removing scheduled task...
schtasks /delete /tn "ClipFix" /f 2>nul
echo [OK] ClipFix service removed.
pause
