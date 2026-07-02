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

# 停掉舊的伺服器（清除 5001–5020 範圍內舊的 process）
for p in $(seq 5001 5020); do
  lsof -ti:$p | xargs kill -9 2>/dev/null
done

# 啟動伺服器
rm -f .port
venv/bin/python3 app.py &
# 等待 app.py 寫出 .port 檔（最多 10 秒）
for i in $(seq 1 10); do
  sleep 1
  [ -f .port ] && break
done

# 讀取實際使用的 port
PORT=$(cat .port 2>/dev/null || echo "5001")

# 取得本機 IP
IP=$(ipconfig getifaddr en0 2>/dev/null || ipconfig getifaddr en1 2>/dev/null || echo "localhost")

echo ""
echo "=============================="
echo "  ✅ 啟動完成！"
echo ""
echo "  本機使用："
echo "  http://localhost:$PORT"
echo ""
echo "  分享給同事（需同網路）："
echo "  http://$IP:$PORT"
echo "=============================="
echo ""

# 複製本機網址到剪貼簿
echo "http://localhost:$PORT" | pbcopy
echo "  (本機網址已複製到剪貼簿)"
echo ""

# 開啟瀏覽器
open "http://localhost:$PORT"

echo "  關閉此視窗即停止伺服器"
wait
