# Excel/Word to Markdown Converter セットアップスクリプト

Write-Host "==================================" -ForegroundColor Cyan
Write-Host "Excel/Word to Markdown Converter" -ForegroundColor Cyan
Write-Host "セットアップスクリプト" -ForegroundColor Cyan
Write-Host "==================================" -ForegroundColor Cyan
Write-Host ""

# Python のバージョンチェック
Write-Host "Pythonのバージョンを確認しています..." -ForegroundColor Yellow
$pythonVersion = python --version 2>&1
Write-Host $pythonVersion -ForegroundColor Green
Write-Host ""

# 仮想環境の作成
Write-Host "仮想環境を作成しています..." -ForegroundColor Yellow
if (Test-Path "venv") {
    Write-Host "既存の仮想環境が見つかりました。スキップします。" -ForegroundColor Green
} else {
    python -m venv venv
    Write-Host "仮想環境を作成しました。" -ForegroundColor Green
}
Write-Host ""

# 仮想環境の有効化
Write-Host "仮想環境を有効化しています..." -ForegroundColor Yellow
& .\venv\Scripts\Activate.ps1

# 依存パッケージのインストール
Write-Host "依存パッケージをインストールしています..." -ForegroundColor Yellow
pip install -r requirements.txt

Write-Host ""
Write-Host "==================================" -ForegroundColor Cyan
Write-Host "セットアップが完了しました！" -ForegroundColor Green
Write-Host "==================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "アプリケーションを起動するには:" -ForegroundColor Yellow
Write-Host "  python converter_app.py" -ForegroundColor White
Write-Host ""

