#!/bin/bash
cd "$(dirname "$0")"
clear
echo "=============================="
echo "  KKday 工單自動化 — 啟動中"
echo "=============================="
echo ""

# 確認有沒有安裝過
if [ ! -d "venv" ]; then
  echo "❌ 尚未安裝！請先雙擊 setup.command 完成安裝"
  read -p "按 Enter 關閉..."
  exit 1
fi

# 自動同步最新版本
if [ -d ".git" ]; then
  echo "⏳ 同步最新版本..."
  git pull origin main --quiet 2>/dev/null && echo "✅ 已是最新版本" || echo "⚠️  無法連線更新（使用本機版本）"
  echo ""
fi

# 停掉舊的伺服器
lsof -ti:5001 | xargs kill -9 2>/dev/null

# 啟動伺服器
source venv/bin/activate
venv/bin/python3 app.py &
sleep 2

# 取得本機 IP
IP=$(ipconfig getifaddr en0 2>/dev/null || ipconfig getifaddr en1 2>/dev/null || echo "localhost")

echo ""
echo "=============================="
echo "  ✅ 啟動完成！"
echo ""
echo "  本機使用："
echo "  http://localhost:5001"
echo ""
echo "  分享給同事（需同網路）："
echo "  http://$IP:5001"
echo "=============================="
echo ""

# 複製本機網址到剪貼簿
echo "http://localhost:5001" | pbcopy
echo "  (本機網址已複製到剪貼簿)"
echo ""

# 開啟瀏覽器
open "http://localhost:5001"

echo "  關閉此視窗即停止伺服器"
wait
