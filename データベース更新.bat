@echo off
chcp 65001 >nul
cd /d %~dp0

echo ==========================================
echo       教材データベース更新ツール v2
echo ==========================================
echo.
echo Excelからデータを直接読み込んでデータベースを更新します。
echo Excelがインストールされている必要があります。
echo.

powershell -NoProfile -ExecutionPolicy Bypass -File "generate_database.ps1"

if %errorlevel% neq 0 (
    echo.
    echo エラー: 更新に失敗しました。
    pause
    exit /b %errorlevel%
)

echo.
echo ==========================================
echo       更新が正常に完了しました！
echo ==========================================
echo.
pause
