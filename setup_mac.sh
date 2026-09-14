#!/usr/bin/env bash
# 一键安装脚本（macOS）
# 作用：检查 Python → 建独立虚拟环境 → 装依赖 → 自检能否启动
set -e
cd "$(dirname "$0")"

echo "== 微信小程序链接批量采集器 · macOS 安装 =="

# ---------- 1. 找一个可用的 python3 ----------
PY=""
for c in python3.11 python3.12 python3.13 python3; do
  if command -v "$c" >/dev/null 2>&1; then PY="$c"; break; fi
done
if [ -z "$PY" ]; then
  echo "❌ 没找到 python3。请先安装 Python 3.9+："
  echo "   推荐从 https://www.python.org/downloads/macos/ 下载安装（.pkg 安装包）"
  exit 1
fi
PY_BIN="$(command -v "$PY")"
echo "✅ 使用 Python: $PY_BIN ($($PY_BIN -V 2>&1))"

# ---------- 2. 关键检查：tkinter 是否可用 ----------
# 图形界面依赖 tkinter。macOS 自带的 /usr/bin/python3（Xcode 命令行工具版）是「非 framework」
# 构建，tkinter 能 import 但窗口一片白屏。必须用 python.org 的 framework 版 Python。
if ! "$PY_BIN" -c "import tkinter" >/dev/null 2>&1; then
  echo "❌ 这个 Python 缺少 tkinter，无法显示界面。"
  echo "   请安装 python.org 版本：https://www.python.org/downloads/macos/"
  exit 1
fi
case "$PY_BIN" in
  /usr/bin/python3*)
    echo "⚠️  你正在用 Xcode 自带的 Python，它的 tkinter 会白屏。"
    echo "   请安装 python.org 版 Python 后重跑本脚本：https://www.python.org/downloads/macos/"
    exit 1
    ;;
esac

# ---------- 3. 建虚拟环境 + 装依赖 ----------
if [ ! -d ".venv" ]; then
  echo "→ 创建虚拟环境 .venv ..."
  "$PY_BIN" -m venv .venv
fi
echo "→ 安装依赖（第一次约 30 秒）..."
./.venv/bin/python -m pip install --upgrade pip -q
./.venv/bin/python -m pip install -r requirements.txt -q

# ---------- 4. 自检 ----------
echo "→ 自检 ..."
./.venv/bin/python -c "
import tkinter, pyautogui, pyperclip, openpyxl
print('✅ tkinter / pyautogui / pyperclip / openpyxl 全部就绪')
"

chmod +x run_mac.command 2>/dev/null || true

cat <<'EOF'

==================== 安装完成 ====================
启动方式：双击 run_mac.command（或命令行 ./run_mac.command）

首次使用还需授权（只需一次）：
  系统设置 → 隐私与安全性 → 辅助功能   → 勾上「终端 / Terminal」
  系统设置 → 隐私与安全性 → 屏幕录制 → 勾上「终端 / Terminal」
（工具本身不截图，但 pyautogui 的部分能力在 macOS 上要求这项授权）
=================================================
EOF
