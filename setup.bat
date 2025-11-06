@echo off
REM Excel/Word to Markdown Converter セットアップスクリプト

echo ==================================
echo Excel/Word to Markdown Converter
echo セットアップスクリプト
echo ==================================
echo.

REM Python のバージョンチェック
echo Pythonのバージョンを確認しています...
python --version
echo.

REM 仮想環境の作成
echo 仮想環境を作成しています...
if exist venv (
    echo 既存の仮想環境が見つかりました。スキップします。
) else (
    python -m venv venv
    echo 仮想環境を作成しました。
)
echo.

REM 仮想環境の有効化
echo 仮想環境を有効化しています...
call venv\Scripts\activate.bat

REM 依存パッケージのインストール
echo 依存パッケージをインストールしています...
pip install -r requirements.txt

echo.
echo ==================================
echo セットアップが完了しました！
echo ==================================
echo.
echo アプリケーションを起動するには:
echo   python converter_app.py
echo.

pause

