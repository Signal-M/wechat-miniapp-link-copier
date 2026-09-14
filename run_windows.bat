@echo off
chcp 65001 >nul
REM 微信小程序链接批量采集器 —— Windows 启动器（双击运行）
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
  echo 还没安装。请先双击 setup_windows.bat 完成安装。
  pause
  exit /b 1
)

".venv\Scripts\python.exe" wechat_scraper_gui.py
if errorlevel 1 (
  echo.
  echo [启动失败] 详情见同目录 wechat_scraper_gui.error.log
  pause
)
