# Excel/Word to Markdown Converter

ExcelファイルとWordファイルをMarkdown形式に変換するGUIアプリケーションです。

## 特徴

- 📁 直感的なGUIインターフェース
- 📊 Excel (.xlsx, .xls) のサポート
- 📝 Word (.docx, .doc) のサポート
- 🔄 PDF経由での高品質な変換
- 🖼️ 画像の自動抽出と埋め込み

## 必要要件

- Windows OS
- Python 3.8以上
- Microsoft Excel（Excelファイル変換用）
- Microsoft Word（Wordファイル変換用）

## インストール手順

### 1. 仮想環境の作成と有効化

```powershell
# 仮想環境を作成
python -m venv venv

# 仮想環境を有効化
.\venv\Scripts\Activate.ps1
```

もしくは、コマンドプロンプトの場合：

```cmd
# 仮想環境を作成
python -m venv venv

# 仮想環境を有効化
venv\Scripts\activate.bat
```

### 2. 依存パッケージのインストール

```powershell
pip install -r requirements.txt
```

## 使い方

### 1. アプリケーションの起動

```powershell
# 仮想環境が有効化されていることを確認
python converter_app.py
```

### 2. ファイルの変換

1. **参照ボタン**をクリックして、変換したいExcelまたはWordファイルを選択
2. **変換ボタン**をクリックして変換を実行
3. 変換が完了すると、元のファイルと同じ場所にMarkdownファイル（.md）が作成されます

## 変換の流れ

```
Excel/Word ファイル
    ↓
   PDF に変換（pywin32使用）
    ↓
Markdown に変換（PyMuPDF使用）
    ↓
.md ファイル出力
```

## サポートされているファイル形式

- Excel: `.xlsx`, `.xls`
- Word: `.docx`, `.doc`

## 出力ファイル

- 元のファイル名に `.md` 拡張子が付いたMarkdownファイル
- ドキュメント内の画像は個別のファイルとして抽出され、Markdownに参照が埋め込まれます

## トラブルシューティング

### pywin32のインストールエラー

```powershell
# pywin32のインストール後、以下を実行
python venv\Scripts\pywin32_postinstall.py -install
```

### COMエラーが発生する場合

Microsoft ExcelまたはWordが正しくインストールされているか確認してください。

### 権限エラー

PowerShellで実行ポリシーエラーが出る場合：

```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

## ライセンス

MIT License

## 開発者向け

### 使用技術

- **GUI**: tkinter
- **Office操作**: pywin32
- **PDF処理**: PyMuPDF (fitz)
- **仮想環境**: venv

### プロジェクト構造

```
convert-markdown/
├── venv/                  # 仮想環境（作成後）
├── converter_app.py       # メインアプリケーション
├── requirements.txt       # 依存パッケージ
├── setup.ps1             # セットアップスクリプト（PowerShell）
└── README.md             # このファイル
```

