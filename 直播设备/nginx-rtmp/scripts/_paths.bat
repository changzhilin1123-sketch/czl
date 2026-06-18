@echo off
set "SCRIPTS_DIR=%~dp0"
for %%I in ("%SCRIPTS_DIR%..") do set "PKG_ROOT=%%~fI"
if exist "%SCRIPTS_DIR%paths.local.bat" call "%SCRIPTS_DIR%paths.local.bat"
if not defined ROOT set "ROOT=%PKG_ROOT%\runtime"
for %%I in ("%ROOT%") do set "ROOT=%%~fI"
set "CONF_SRC=%PKG_ROOT%\conf\nginx.conf"
set "ZIP=%ROOT%\nginx-rtmp.zip"
set "DOWNLOAD_URL=https://github.com/iliweii/nginx-rtmp-win64/releases/download/v1.0.0/nginx-rtmp-win64.zip"
