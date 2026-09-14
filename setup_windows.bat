@echo off
chcp 65001 >nul
REM ===============================================
REM  微信小程序链接批量采集器 —— Windows 一键安装
REM ===============================================
cd /d "%~dp0"

echo == 微信小程序链接批量采集器 . Windows 安装 ==

where python >nul 2>nul
if errorlevel 1 (
  echo [X] 没找到 python。请先安装 Python 3.9+：
  echo     推荐从 https://www.python.org/downloads/windows/ 下载，
  echo     安装时务必勾选 "Add python.exe to PATH"。
  pause
  exit /b 1
)

python -c "import tkinter" >nul 2>nul
if errorlevel 1 (
  echo [X] 这个 Python 缺少 tkinter，无法显示界面。请重装官方 python.org 版本。
  pause
  exit /b 1
)

if not exist ".venv" (
  echo -^> 创建虚拟环境 .venv ...
  python -m venv .venv
)

echo -^> 安装依赖（第一次约 30 秒）...
".venv\Scripts\python.exe" -m pip install --upgrade pip -q
".venv\Scripts\python.exe" -m pip install -r requirements.txt -q

echo -^> 自检 ...
".venv\Scripts\python.exe" -c "import tkinter, pyautogui, pyperclip, openpyxl; print('[OK] tkinter / pyautogui / pyperclip / openpyxl 全部就绪')"

echo.
echo ==================== 安装完成 ====================
echo 启动方式：双击 run_windows.bat
echo.
echo Windows 无需额外授权，但请把微信窗口固定在屏幕上的固定位置，
echo 运行期间不要移动、最小化或改变缩放比例。
echo =================================================
echo.
pause
