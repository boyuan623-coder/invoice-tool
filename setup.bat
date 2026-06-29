@echo off
setlocal enabledelayedexpansion
title 发票识别工具 - 安装

:: ============================================
:: 第一步：检查 Python
:: ============================================
echo.
echo ========================================
echo   发票识别工具 - 环境安装
echo ========================================
echo.
echo [1/3] 检查 Python 环境...

set "PYTHON_OK=0"

:: 尝试 python3
python3 --version >nul 2>&1
if !errorlevel! equ 0 (
    set "PYTHON_CMD=python3"
    set "PYTHON_OK=1"
)

:: 尝试 python
if !PYTHON_OK! equ 0 (
    python --version >nul 2>&1
    if !errorlevel! equ 0 (
        set "PYTHON_CMD=python"
        set "PYTHON_OK=1"
    )
)

:: 尝试 py 启动器
if !PYTHON_OK! equ 0 (
    py --version >nul 2>&1
    if !errorlevel! equ 0 (
        set "PYTHON_CMD=py"
        set "PYTHON_OK=1"
    )
)

if !PYTHON_OK! equ 1 (
    echo       已找到: !PYTHON_CMD!
    goto :install_deps
)

:: ============================================
:: 第二步：没 Python，自动安装
:: ============================================
echo       未检测到 Python，开始自动安装...
echo.

set "PYTHON_INSTALLER=%TEMP%\python-installer.exe"
set "DOWNLOAD_URL=https://www.python.org/ftp/python/3.12.0/python-3.12.0-amd64.exe"

echo       正在下载 Python 3.12 （约 25MB）...
echo       如果下载失败请手动访问: https://www.python.org/downloads/
echo.

:: 用 PowerShell 下载（兼容大多数 Win10+）
powershell -Command "[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; (New-Object Net.WebClient).DownloadFile('%DOWNLOAD_URL%', '%PYTHON_INSTALLER%')" 2>nul

if not exist "%PYTHON_INSTALLER%" (
    echo.
    echo ========================================
    echo   下载失败！
    echo ========================================
    echo.
    echo   可能原因：网络不通或被防火墙拦截
    echo.
    echo   请手动操作：
    echo   1. 打开浏览器访问
    echo      https://www.python.org/downloads/
    echo   2. 下载 Python 3.10 以上版本
    echo   3. 安装时务必勾选 "Add Python to PATH"
    echo   4. 安装完成后重新双击 setup.bat
    echo.
    pause
    exit /b 1
)

echo       正在安装 Python （静默安装，请稍候）...
start /wait "" "%PYTHON_INSTALLER%" /quiet InstallAllUsers=1 PrependPath=1 Include_test=0 2>nul

:: 等待安装完成
timeout /t 5 /nobreak >nul

:: 清理安装包
del "%PYTHON_INSTALLER%" 2>nul

:: 刷新环境变量
call :refresh_env

:: 再次检查
python --version >nul 2>&1
if !errorlevel! equ 0 (
    set "PYTHON_CMD=python"
    echo       Python 安装成功！
    echo.
    echo       请关闭此窗口，重新双击 setup.bat 继续安装依赖
    pause
    exit /b 0
)

python3 --version >nul 2>&1
if !errorlevel! equ 0 (
    set "PYTHON_CMD=python3"
    echo       Python 安装成功！
    echo       请关闭此窗口，重新双击 setup.bat
    pause
    exit /b 0
)

echo.
echo ========================================
echo   安装可能未生效
echo ========================================
echo.
echo   请重新启动电脑后再双击 setup.bat
echo   或手动安装 Python:
echo   https://www.python.org/downloads/
echo   （安装时勾选 Add to PATH）
echo.
pause
exit /b 1


:: ============================================
:: 第三步：有 Python，装依赖
:: ============================================
:install_deps
echo.
echo [2/3] 安装依赖包（首次约 5-10 分钟）...
echo.

set "REQ_FILE=%~dp0requirements.txt"
if not exist "%REQ_FILE%" (
    echo       错误：找不到 requirements.txt
    echo       请确保所有文件完整
    pause
    exit /b 1
)

:: 先升级 pip
echo       升级 pip...
!PYTHON_CMD! -m pip install --upgrade pip -q 2>nul

:: 安装 CPU 版 Paddle（兼容无 GPU 电脑）
echo       安装 PaddlePaddle CPU...
!PYTHON_CMD! -m pip install paddlepaddle==3.1.0 -i https://pypi.tuna.tsinghua.edu.cn/simple
if !errorlevel! neq 0 (
    !PYTHON_CMD! -m pip install paddlepaddle==3.1.0
)

:: 安装依赖，先用国内镜像
echo       安装依赖包...
!PYTHON_CMD! -m pip install -r "%REQ_FILE%" -i https://pypi.tuna.tsinghua.edu.cn/simple

if !errorlevel! neq 0 (
    echo.
    echo       国内镜像失败，切换默认源重试...
    !PYTHON_CMD! -m pip install -r "%REQ_FILE%"
    if !errorlevel! neq 0 (
        echo.
        echo ========================================
        echo   依赖安装失败
        echo ========================================
        echo.
        echo   可能原因：网络不通
        echo   请检查网络后重新双击 setup.bat
        echo.
        pause
        exit /b 1
    )
)

echo.
echo [3/3] 完成！
echo.
echo ========================================
echo   安装成功！
echo ========================================
echo.
echo   现在可以双击 "invoice_tool.bat" 使用了
echo   浏览器会自动打开，上传 PDF 即可识别
echo.
echo   或直接运行: python app.py
echo.
echo   如果双击没反应：
echo   在本文件夹按住 Shift + 右键
echo   选择 "在此处打开 PowerShell"
echo   输入 python app.py 查看错误信息
echo ========================================
echo.
pause
exit /b 0


:: ============================================
:: 辅助：刷新环境变量
:: ============================================
:refresh_env
:: 从注册表读取最新的 PATH
for /f "tokens=2*" %%a in ('reg query "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment" /v PATH 2^>nul') do set "NEW_PATH=%%b"
if defined NEW_PATH set "PATH=%NEW_PATH%"
for /f "tokens=2*" %%a in ('reg query "HKCU\Environment" /v PATH 2^>nul') do set "NEW_PATH=%%b"
if defined NEW_PATH set "PATH=%PATH%;%NEW_PATH%"
goto :eof
