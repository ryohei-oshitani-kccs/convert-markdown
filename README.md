# Excel/Word → Markdown 変換ツール

ExcelファイルまたはWordファイルをMarkdown形式に変換するGUIアプリケーションです。

## 機能

- 📁 **ファイル選択**: 参照ボタンから変換したいExcel/Wordファイルを選択
- 🔄 **自動変換**: 選択したファイルをPDF経由でMarkdownに変換
- 💾 **同じ場所に保存**: 変換されたMarkdownファイルは元のファイルと同じディレクトリに保存
- ✅ **対応形式**: `.xlsx`, `.xls`, `.docx`, `.doc`

## 変換フロー

```
Excel/Word → PDF → Markdown
```

1. 選択されたファイル(Excel/Word)をPDF形式に変換
2. 生成されたPDFをMarkdown形式に変換
3. 元のファイルと同じ場所に`.md`ファイルを出力

## 必要な環境

### システム要件

- **Python 3.7以上**
- **Windows OS** (pywin32はWindows専用)
- **Microsoft Office** (Excel/Word)がインストールされている必要があります

### Pythonパッケージのインストール

```bash
pip install -r requirements.txt
```

## 使い方

1. プログラムを起動:
```bash
python converter.py
```

2. **参照ボタン**をクリックして、変換したいExcelまたはWordファイルを選択

3. **変換ボタン**をクリックして変換を実行

4. 変換が完了すると、選択したファイルと同じディレクトリに`.md`ファイルが作成されます

## プロジェクト構成

```
convert-markdown/
├── converter.py         # メインアプリケーション
├── requirements.txt     # 必要なPythonパッケージ
├── README.md           # このファイル
└── LICENSE             # ライセンス
```

## 注意事項

- **Windows専用**: pywin32ライブラリを使用しているため、WindowsOSでのみ動作します
- **Microsoft Officeが必要**: Excel/Wordがインストールされていない場合、PDFへの変換に失敗します
- 複雑なレイアウトや画像を含むファイルの場合、Markdown変換の結果が期待通りにならない場合があります
- 大きなファイルの場合、変換に時間がかかる場合があります
- 変換中はExcel/Wordが一時的にバックグラウンドで起動しますが、自動的に終了します

## トラブルシューティング

### "pywin32がインストールされていません"エラー

pywin32をインストールしてください:
```bash
pip install pywin32
```

インストール後、以下のコマンドを実行してください:
```bash
python Scripts/pywin32_postinstall.py -install
```

### "Excel PDF変換エラー" または "Word PDF変換エラー"

- Microsoft Office (Excel/Word)がインストールされているか確認してください
- Officeライセンスがアクティベートされているか確認してください
- ファイルが他のプログラムで開かれていないか確認してください

### "PyMuPDFがインストールされていません"エラー

必要なパッケージをインストールしてください:
```bash
pip install PyMuPDF
```

## Windows実行ファイル（exe）のビルド

Linux環境で開発している場合、pywin32のインストールができません。
Windows環境で動作する実行ファイル（exe）を作成する手順については、
`BUILD_INSTRUCTIONS.md` を参照してください。

### クイックスタート（Windows環境で実行）

```bash
# 1. 必要なパッケージをインストール
pip install -r requirements_build.txt

# 2. exeファイルをビルド
python build_exe.py

# 3. 実行ファイルが生成されます
# dist/MarkdownConverter.exe
```

詳細な手順とトラブルシューティングは `BUILD_INSTRUCTIONS.md` をご覧ください。

## ライセンス

このプロジェクトのライセンスについては、LICENSEファイルを参照してください。

