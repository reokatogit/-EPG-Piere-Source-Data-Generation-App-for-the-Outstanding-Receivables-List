@echo off
chcp 65001 > nul
echo ============================================
echo  Piere未収リスト元データ作成ツール ビルド
echo ============================================
echo.

REM 依存ライブラリのインストール
echo [1/2] 依存ライブラリをインストールしています...
pip install -r requirements.txt
if errorlevel 1 (
    echo エラー: pip install に失敗しました。
    pause
    exit /b 1
)

echo.
echo [2/2] exe ファイルをビルドしています...
pyinstaller ^
    --onefile ^
    --windowed ^
    --name "Piere未収リスト元データ作成" ^
    --add-data "utils.py;." ^
    --add-data "processor.py;." ^
    --add-data "gui.py;." ^
    --add-data "fast_xlsx_reader.py;." ^
    main.py

if errorlevel 1 (
    echo エラー: ビルドに失敗しました。
    pause
    exit /b 1
)

echo.
echo ============================================
echo  ビルド完了！
echo  dist\Piere未収リスト元データ作成.exe
echo ============================================
pause
