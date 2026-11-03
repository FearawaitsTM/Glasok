@echo off
setlocal
powershell.exe -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "%~dp0start_glazok.ps1"
if not errorlevel 1 exit /b 0
echo.
echo Glasok could not start. Read the message above.
pause
exit /b 1
