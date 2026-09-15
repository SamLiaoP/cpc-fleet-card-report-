@echo off
chcp 65001 >nul
REM 切換到此啟動檔所在的資料夾
cd /d "%~dp0"

echo ================================
echo  中油車隊卡 交易彙整工具
echo ================================
echo.

echo 從 GitHub 拉取最新版本...
git pull --ff-only
if %errorlevel%==0 (
    echo 已更新到最新版本
) else (
    echo git pull 失敗（可能無網路或有衝突），將使用本地版本繼續執行
)
echo.

echo 執行 process_transactions.py
echo --------------------------------

REM 嘗試用 python，失敗則改用 py 啟動器
where python >nul 2>nul
if %errorlevel%==0 (
    python process_transactions.py
) else (
    py process_transactions.py
)

echo --------------------------------
echo.
pause
