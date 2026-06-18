@echo off
setlocal
call "%~dp0_paths.bat"

echo 安装目录: %ROOT%
if not exist "%ROOT%" mkdir "%ROOT%"

if exist "%ROOT%\nginx.exe" (
    echo [SKIP] nginx.exe 已存在
    goto :copy_conf
)

echo 正在下载 nginx-rtmp-win64 ...
powershell -NoProfile -ExecutionPolicy Bypass -Command "Invoke-WebRequest -Uri '%DOWNLOAD_URL%' -OutFile '%ZIP%' -UseBasicParsing"
if errorlevel 1 (
    echo [ERROR] 下载失败，请浏览器打开:
    echo %DOWNLOAD_URL%
    exit /b 1
)

echo 正在解压 ...
powershell -NoProfile -ExecutionPolicy Bypass -Command "Expand-Archive -Path '%ZIP%' -DestinationPath '%ROOT%' -Force"
if not exist "%ROOT%\nginx.exe" (
    for /d %%d in ("%ROOT%\*") do (
        if exist "%%d\nginx.exe" xcopy /E /Y "%%d\*" "%ROOT%\" >nul
    )
)

:copy_conf
if not exist "%ROOT%\conf" mkdir "%ROOT%\conf"
if not exist "%ROOT%\logs" mkdir "%ROOT%\logs"
if not exist "%ROOT%\temp" mkdir "%ROOT%\temp"
copy /Y "%CONF_SRC%" "%ROOT%\conf\nginx.conf" >nul
echo [OK] 配置已同步到 %ROOT%\conf\nginx.conf

powershell -NoProfile -ExecutionPolicy Bypass -Command "New-NetFirewallRule -DisplayName 'Nginx-RTMP-1935-In' -Direction Inbound -Protocol TCP -LocalPort 1935 -Action Allow -Profile Private,Domain -ErrorAction SilentlyContinue | Out-Null; New-NetFirewallRule -DisplayName 'Nginx-HTTP-8080-In' -Direction Inbound -Protocol TCP -LocalPort 8080 -Action Allow -Profile Private,Domain -ErrorAction SilentlyContinue | Out-Null"

echo.
echo 安装完成。下一步: 2-启动.bat  然后  3-验证.bat
endlocal
