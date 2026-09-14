#!/bin/bash
# 微信小程序链接批量采集器 —— macOS 启动器
# 双击本文件即可运行（若提示无法打开，先执行：chmod +x run_mac.command）
cd "$(dirname "$0")" || exit 1

if [ ! -x ".venv/bin/python3" ]; then
  echo "还没安装。请先双击 setup_mac.sh 完成安装（或执行 ./setup_mac.sh）。"
  read -n 1 -s -r -p "按任意键关闭窗口…"
  exit 1
fi

".venv/bin/python3" wechat_scraper_gui.py
ec=$?
if [ "$ec" -ne 0 ]; then
  echo ""
  echo "[启动失败] 退出码 $ec，详情见同目录 wechat_scraper_gui.error.log"
  read -n 1 -s -r -p "按任意键关闭窗口…"
fi
