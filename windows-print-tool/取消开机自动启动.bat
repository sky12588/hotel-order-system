@echo off
chcp 65001 >nul
set "SHORTCUT=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup\酒店打印助手.lnk"
if exist "%SHORTCUT%" del "%SHORTCUT%"
echo 已取消开机自动启动酒店打印助手。
pause
