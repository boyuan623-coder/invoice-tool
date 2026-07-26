@echo off
setlocal EnableExtensions
cd /d "%~dp0"
title Invoice Tool Share

set "PYCMD="
python --version >nul 2>&1 && set "PYCMD=python"
if "%PYCMD%"=="" (
    py --version >nul 2>&1 && set "PYCMD=py"
)
if "%PYCMD%"=="" (
    python3 --version >nul 2>&1 && set "PYCMD=python3"
)

if "%PYCMD%"=="" (
    echo.
    echo ========================================
    echo   Python not found. Run setup.cmd first.
    echo ========================================
    echo.
    pause
    exit /b 1
)

echo.
echo ========================================
echo   Invoice Tool - Share Mode
echo   Local:  http://127.0.0.1:5000
echo   Public link will appear in 10-30 seconds.
echo   Close this window to stop sharing.
echo ========================================
echo.

"%PYCMD%" "%~dp0app.py" --share
set "EXIT_CODE=%ERRORLEVEL%"

echo.
if not "%EXIT_CODE%"=="0" goto FAIL
echo   Service stopped. Public link is invalid now.
goto END

:FAIL
echo ========================================
echo   Start failed. Check errors above.
echo   Tips: port 5000 busy, or network issue.
echo   Try setup.cmd then run again.
echo ========================================

:END
echo.
pause
exit /b %EXIT_CODE%
