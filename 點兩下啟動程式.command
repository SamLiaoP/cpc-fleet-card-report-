#!/bin/bash
# 切換到此啟動檔所在的資料夾
cd "$(dirname "$0")" || exit 1

echo "================================"
echo " 中油車隊卡 交易彙整工具"
echo "================================"
echo ""

echo "🔄 從 GitHub 拉取最新版本..."
if git pull --ff-only; then
    echo "✅ 已更新到最新版本"
else
    echo "⚠️  git pull 失敗（可能無網路或有衝突），將使用本地版本繼續執行"
fi
echo ""

echo "▶️  執行 process_transactions.py"
echo "--------------------------------"
python3 process_transactions.py
echo "--------------------------------"
echo ""

read -n 1 -s -r -p "按任意鍵關閉視窗..."
echo ""
