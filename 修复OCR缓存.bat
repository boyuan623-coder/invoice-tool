@echo off
chcp 65001 >nul
echo ========================================
echo   OCR 缓存修复工具
echo   用于解决 PaddleOCR 初始化失败问题
echo ========================================
echo.

set "DELETED=0"

call :delete_dir "%USERPROFILE%\.paddlex"
call :delete_dir "%USERPROFILE%\.paddleocr"

:: 中文用户名下的缓存重定向目录（优先 D/E/F/C 盘）
for %%D in (D E F C) do (
    if exist "%%D:\" (
        call :delete_dir "%%D:\paddle_ocr_cache\.paddlex"
        call :delete_dir "%%D:\paddle_ocr_cache\.paddleocr"
    )
)

echo.
echo ========================================
echo   缓存清理完成！
echo   下次运行发票识别工具时会重新下载模型
echo   （需要联网，约 500MB，首次较慢）
echo   无需重启电脑，直接重新上传 PDF 即可
echo ========================================
pause
exit /b 0

:delete_dir
if not exist "%~1" (
    echo [跳过] 目录不存在: %~1
    goto :eof
)
echo [清理] 正在删除: %~1
rmdir /s /q "%~1" 2>nul
if exist "%~1" (
    echo [失败] 无法删除，请手动删除: %~1
) else (
    echo [完成] 已删除
    set "DELETED=1"
)
goto :eof
