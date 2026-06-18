@echo off
setlocal
call "%~dp0_paths.bat"

if not exist "%ROOT%\nginx.exe" (
    echo [ERROR] 未找到 %ROOT%\nginx.exe
    echo 请先运行 1-安装.bat 或 install-nginx-rtmp.bat
    exit /b 1
)

if not exist "%ROOT%\conf" mkdir "%ROOT%\conf"
if not exist "%ROOT%\logs" mkdir "%ROOT%\logs"
if not exist "%ROOT%\temp" mkdir "%ROOT%\temp"
copy /Y "%CONF_SRC%" "%ROOT%\conf\nginx.conf" >nul

cd /d "%ROOT%"
nginx.exe -p "%ROOT%" -c conf/nginx.conf -t
if errorlevel 1 exit /b 1

tasklist /FI "IMAGENAME eq nginx.exe" 2>nul | find /I "nginx.exe" >nul
if %errorlevel%==0 (
    nginx.exe -p "%ROOT%" -c conf/nginx.conf -s reload
    echo [OK] Nginx 已在运行，已 reload 配置
) else (
    start "" /B nginx.exe -p "%ROOT%" -c conf/nginx.conf
    timeout /t 1 /nobreak >nul
    echo [OK] Nginx RTMP 已启动
)

echo.
echo 安装目录: %ROOT%
echo 推流/拉流: rtmp://^<本机IP^>:1935/live/drone
echo 本机测试:  rtmp://127.0.0.1:1935/live/drone
echo 状态页:    http://127.0.0.1:8080/stat
echo.
for /f "tokens=*" %%i in ('powershell -NoProfile -Command "Get-NetIPAddress -AddressFamily IPv4 | Where-Object { $_.IPAddress -notlike '127.*' -and $_.PrefixOrigin -ne 'WellKnown' } | ForEach-Object { $_.InterfaceAlias + ': ' + $_.IPAddress }"') do echo   %%i
echo.
netstat -ano | findstr "LISTENING" | findstr ":1935 :8080"
endlocal
