@echo off
setlocal
call "%~dp0_paths.bat"
cd /d "%ROOT%" 2>nul
if exist "%ROOT%\nginx.exe" nginx.exe -p "%ROOT%" -c conf/nginx.conf -s stop 2>nul
timeout /t 1 /nobreak >nul
taskkill /F /IM nginx.exe 2>nul
echo [OK] Nginx 已停止
endlocal
