@echo off
setlocal EnableExtensions
cd /d "%~dp0"
title Invoice Tool Fixed Share

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

if not exist "%~dp0tools\tunnel-config.yml" (
    echo.
    echo Fixed tunnel not configured yet.
    echo Please run setup_fixed_tunnel.bat first.
    echo.
    pause
    exit /b 1
)

echo.
echo ========================================
echo   Invoice Tool - Fixed Share
echo   URL: https://invoice.bingshanforprivate.asia
echo   Close this window to stop.
echo ========================================
echo.

:: 本机有 NVIDIA GPU 时优先使用 GPU 加速 OCR
set "INVOICE_TOOL_DEVICE=gpu:0"
"%PYCMD%" "%~dp0app.py" --share-fixed
set "EXIT_CODE=%ERRORLEVEL%"

echo.
if not "%EXIT_CODE%"=="0" goto FAIL
echo   Service stopped.
goto END

:FAIL
echo   Start failed. Check errors above.
echo   Tip: run setup_fixed_tunnel.bat if not done.

:END
echo.
pause
exit /b %EXIT_CODE%
