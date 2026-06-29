@echo off
chcp 65001 >nul
setlocal
title 电子发票识别工具

:: 始终切换到 bat 所在目录（路径含中文/空格也适用）
cd /d "%~dp0"

:: 查找 Python
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
    echo   未找到 Python，请先运行 setup.cmd
    echo ========================================
    echo.
    pause
    exit /b 1
)

echo.
echo ========================================
echo   电子发票识别工具
echo   浏览器将自动打开 http://127.0.0.1:5000
echo   关闭此窗口即可停止服务
echo ========================================
echo.

:: 在本窗口直接启动（可看到日志；app.py 会自动打开浏览器）
"%PYCMD%" "%~dp0app.py"
set "EXIT_CODE=%ERRORLEVEL%"

echo.
if not "%EXIT_CODE%"=="0" (
    echo ========================================
    echo   启动失败，请检查上方错误信息
    echo   常见原因：端口 5000 被占用 / 依赖未安装
    echo   可尝试运行 setup.cmd 重新安装依赖
    echo ========================================
) else (
    echo   服务已停止
)
echo.
pause
exit /b %EXIT_CODE%
