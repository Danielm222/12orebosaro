@echo off
setlocal
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0reset_torneo.ps1"
echo.
pause
endlocal
