@echo off
setlocal
cd /d "%~dp0.."

set "PORT=8765"

powershell -NoProfile -Command ^
  "$port=%PORT%; $listening=Get-NetTCPConnection -LocalPort $port -State Listen -ErrorAction SilentlyContinue; if(-not $listening){Start-Process -FilePath python -ArgumentList '-m','http.server','%PORT%','--bind','127.0.0.1' -WorkingDirectory '%~dp0..' -WindowStyle Hidden}"

timeout /t 2 /nobreak >nul
start "" "http://127.0.0.1:%PORT%/schermo.html"

endlocal
