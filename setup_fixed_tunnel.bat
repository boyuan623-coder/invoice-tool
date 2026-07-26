@echo off
setlocal EnableExtensions
cd /d "%~dp0"
title Setup Fixed Tunnel

set "PYCMD="
python --version >nul 2>&1 && set "PYCMD=python"
if "%PYCMD%"=="" (
    py --version >nul 2>&1 && set "PYCMD=py"
)
if "%PYCMD%"=="" (
    python3 --version >nul 2>&1 && set "PYCMD=python3"
)

if "%PYCMD%"=="" (
    echo Python not found. Run setup.cmd first.
    pause
    exit /b 1
)

echo.
echo ========================================
echo  Setup fixed URL for:
echo  https://invoice.bingshanforprivate.asia
echo ========================================
echo.
echo BEFORE CONTINUE:
echo  1. Add domain bingshanforprivate.asia to Cloudflare
echo  2. Change nameservers at your domain registrar
echo  3. Wait until DNS works
echo.
echo A browser will open for Cloudflare login.
echo.

"%PYCMD%" "%~dp0tunnel_fixed.py" setup invoice.bingshanforprivate.asia
echo.
pause
exit /b %ERRORLEVEL%
