@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo 正在解除 Windows 对本文件夹脚本的安全限制...
powershell -NoProfile -ExecutionPolicy Bypass -Command "Get-ChildItem -LiteralPath '%~dp0' -Recurse | Unblock-File"
echo.
echo 已处理完成。
echo 现在请双击“启动酒店打印助手.bat”。
echo.
pause
