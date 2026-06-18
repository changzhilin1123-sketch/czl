@echo off
setlocal EnableDelayedExpansion
call "%~dp0_paths.bat"
set "PASS=0"
set "FAIL=0"

echo ========================================
echo   Nginx-RTMP 验证  %date% %time%
echo   目录: %ROOT%
echo ========================================
echo.

call :check "nginx.exe 存在" exist "%ROOT%\nginx.exe"
call :check "nginx 进程运行" process nginx.exe
call :check "端口 1935 监听" port 1935
call :check "端口 8080 监听" port 8080
call :check "状态页 HTTP 200" http

echo ----------------------------------------
echo 本机 IP（RC2 推流用，勿用 127.0.0.1）:
powershell -NoProfile -Command "Get-NetIPAddress -AddressFamily IPv4 | Where-Object { $_.IPAddress -notlike '127.*' -and $_.PrefixOrigin -ne 'WellKnown' } | ForEach-Object { Write-Host ('  ' + $_.InterfaceAlias + ': rtmp://' + $_.IPAddress + ':1935/live/drone') }"
echo.
echo 结果: PASS=!PASS!  FAIL=!FAIL!
if !FAIL! GTR 0 (
    echo [未通过] 请运行 2-启动.bat
    exit /b 1
)
echo [全部通过] 可进行 RC2 推流与 VLC/伴侣窗口采集
exit /b 0

:check
set "LABEL=%~1"
set "TYPE=%~2"
set "ARG=%~3"
set "OK=0"
if "%TYPE%"=="exist" if exist "%ARG%" set "OK=1"
if "%TYPE%"=="process" tasklist /FI "IMAGENAME eq %ARG%" 2>nul | find /I "%ARG%" >nul && set "OK=1"
if "%TYPE%"=="port" netstat -ano | findstr "LISTENING" | findstr ":%ARG% " >nul && set "OK=1"
if "%TYPE%"=="http" (
    powershell -NoProfile -Command "try { $r = Invoke-WebRequest -Uri 'http://127.0.0.1:8080/stat' -UseBasicParsing -TimeoutSec 5; if ($r.StatusCode -eq 200) { exit 0 } else { exit 1 } } catch { exit 1 }" >nul 2>&1
    if !errorlevel!==0 set "OK=1"
)
if "!OK!"=="1" (echo [PASS] %LABEL% & set /a PASS+=1) else (echo [FAIL] %LABEL% & set /a FAIL+=1)
goto :eof
