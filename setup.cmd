@echo off
setlocal
title Invoice Extractor Setup

echo.
echo ========================================
echo   Invoice Extractor - Setup
echo ========================================
echo.
echo [1/3] Checking Python...

set "PYCMD="

python3 --version >nul 2>&1
if %errorlevel% equ 0 set "PYCMD=python3"

if "%PYCMD%"=="" (
    python --version >nul 2>&1
    if %errorlevel% equ 0 set "PYCMD=python"
)

if "%PYCMD%"=="" (
    py --version >nul 2>&1
    if %errorlevel% equ 0 set "PYCMD=py"
)

if not "%PYCMD%"=="" (
    echo        Found: %PYCMD%
    goto :install
)

echo        Python not found, downloading...
echo.
set "INSTALLER=%TEMP%\python-installer.exe"
set "URL=https://www.python.org/ftp/python/3.12.0/python-3.12.0-amd64.exe"

powershell -Command "[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; (New-Object Net.WebClient).DownloadFile('%URL%', '%INSTALLER%')" 2>nul

if not exist "%INSTALLER%" (
    echo.
    echo ========================================
    echo   Download failed!
    echo ========================================
    echo.
    echo   Please install Python manually:
    echo   1. Visit https://www.python.org/downloads/
    echo   2. Download Python 3.10+
    echo   3. Install - CHECK "Add Python to PATH"
    echo   4. Re-run setup.cmd
    echo.
    pause
    exit /b 1
)

echo        Installing Python (please wait)...
start /wait "" "%INSTALLER%" /quiet InstallAllUsers=1 PrependPath=1 Include_test=0
del "%INSTALLER%" 2>nul

:: Refresh PATH from registry
for /f "tokens=2*" %%a in ('reg query "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment" /v PATH 2^>nul') do set "PATH=%%b;%PATH%"
for /f "tokens=2*" %%a in ('reg query "HKCU\Environment" /v PATH 2^>nul') do set "PATH=%%b;%PATH%"

:: Try again
python --version >nul 2>&1
if %errorlevel% equ 0 set "PYCMD=python"

if "%PYCMD%"=="" (
    echo.
    echo ========================================
    echo   Installation may need reboot
    echo ========================================
    echo.
    echo   Please RESTART your computer,
    echo   then double-click setup.cmd again.
    echo.
    echo   Or install manually:
    echo   https://www.python.org/downloads/
    echo   (CHECK "Add Python to PATH")
    echo.
    pause
    exit /b 1
)

echo        Python installed!
echo        Please close this window and re-run setup.cmd
pause
exit /b 0

:install
echo.
echo [2/3] Installing dependencies (5-10 min)...

%PYCMD% -m pip install --upgrade pip -q 2>nul

echo        Installing PaddlePaddle CPU (compatible with all PCs)...
%PYCMD% -m pip install paddlepaddle==3.1.0 -i https://pypi.tuna.tsinghua.edu.cn/simple
if %errorlevel% neq 0 (
    %PYCMD% -m pip install paddlepaddle==3.1.0
)

%PYCMD% -m pip install -r "%~dp0requirements.txt" -i https://pypi.tuna.tsinghua.edu.cn/simple

if %errorlevel% neq 0 (
    echo.
    echo        Mirror failed, trying default...
    %PYCMD% -m pip install -r "%~dp0requirements.txt"
    if %errorlevel% neq 0 (
        echo.
        echo ========================================
        echo   Install failed - check network
        echo ========================================
        pause
        exit /b 1
    )
)

echo.
echo [3/3] Done!
echo.
echo ========================================
echo   Setup complete!
echo ========================================
echo.
echo   Now double-click: invoice_tool.bat
echo   Browser will open automatically.
echo.
echo   Or run manually: python app.py
echo.
echo   If startup fails:
echo   Shift+Right-click here -> Open PowerShell
echo   Type: python app.py
echo ========================================
pause
