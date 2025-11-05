# Windows 実行ファイル (exe) ビルド手順

このドキュメントは、Linux環境で開発したアプリケーションをWindows環境でexeファイルとしてビルドする手順を説明します。

## 重要な注意事項

**Linux環境ではWindows用のexeファイルをビルドできません。**
必ずWindows環境で以下の手順を実行してください。

## 前提条件

- **Windows OS** (Windows 10/11 推奨)
- **Python 3.7以上** がインストールされていること
- **Microsoft Office** (Excel/Word) がインストールされていること

## ビルド手順

### 1. プロジェクトファイルをWindows環境に転送

Linux環境からWindows環境へ、以下のファイルを転送してください：

```
convert-markdown/
├── converter.py
├── build_exe.py
├── requirements_build.txt
└── BUILD_INSTRUCTIONS.md (このファイル)
```

転送方法の例：
- Git リポジトリ経由
- USB メモリ
- ネットワーク共有
- クラウドストレージ (Google Drive, Dropbox など)

### 2. Python仮想環境の作成（推奨）

Windows環境でコマンドプロンプトまたはPowerShellを開き、プロジェクトディレクトリに移動します：

```cmd
cd path\to\convert-markdown
```

仮想環境を作成：

```cmd
python -m venv venv
```

仮想環境を有効化：

**コマンドプロンプトの場合:**
```cmd
venv\Scripts\activate.bat
```

**PowerShellの場合:**
```powershell
venv\Scripts\Activate.ps1
```

※ PowerShellで実行ポリシーエラーが出る場合：
```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

### 3. 必要なパッケージのインストール

```cmd
pip install -r requirements_build.txt
```

pywin32のインストール後、以下も実行：

```cmd
python venv\Scripts\pywin32_postinstall.py -install
```

### 4. exeファイルのビルド

ビルドスクリプトを実行：

```cmd
python build_exe.py
```

ビルドには数分かかる場合があります。

### 5. ビルド結果の確認

ビルドが成功すると、以下のディレクトリに実行ファイルが生成されます：

```
dist/MarkdownConverter.exe
```

## 実行ファイルのテスト

1. `dist` フォルダ内の `MarkdownConverter.exe` をダブルクリックして起動
2. GUIが表示されることを確認
3. サンプルのExcel/Wordファイルで変換をテスト

## 配布

`dist/MarkdownConverter.exe` ファイルを他のWindows環境にコピーすれば、
Pythonがインストールされていない環境でも実行できます。

**注意：**
- 実行環境にはMicrosoft Office (Excel/Word)が必要です
- 初回起動時にWindowsのセキュリティ警告が表示される場合があります

## トラブルシューティング

### ビルドエラー: "PyInstaller not found"

```cmd
pip install pyinstaller
```

### ビルドエラー: "pywin32 import error"

```cmd
pip uninstall pywin32
pip install pywin32
python venv\Scripts\pywin32_postinstall.py -install
```

### exeファイルが起動しない

1. コンソール版でビルドして詳細なエラーを確認：
   - `build_exe.py` 内の `'--windowed',` をコメントアウト
   - 再ビルド

2. または、直接PyInstallerでビルド：
```cmd
pyinstaller --onefile --console converter.py
```

### exeファイルのサイズが大きい

- `--onefile` を `--onedir` に変更すると、複数ファイルに分割されます
- サイズは大きくなりますが、起動が速くなる場合があります

```cmd
pyinstaller --onedir --windowed converter.py
```

## 高度な設定

### アイコンの追加

1. `.ico` 形式のアイコンファイルを用意
2. `build_exe.py` 内の `icon_file = None` を以下のように変更：
```python
icon_file = "icon.ico"
```

### ビルドオプションのカスタマイズ

`build_exe.py` 内の `pyinstaller_args` リストを編集して、
PyInstallerのオプションをカスタマイズできます。

詳細は PyInstaller のドキュメントを参照：
https://pyinstaller.org/en/stable/

## 参考情報

- **PyInstaller 公式ドキュメント**: https://pyinstaller.org/
- **pywin32**: https://github.com/mhammond/pywin32
- **PyMuPDF**: https://pymupdf.readthedocs.io/

## Linux環境で開発を続ける場合

Linux環境では以下のワークフローをお勧めします：

1. Linux環境でコードを開発・編集
2. Git等でバージョン管理
3. Windows環境にチェックアウト
4. Windows環境でビルド・テスト

または、Windows仮想マシン (VirtualBox, VMware等) やWSL2を使用することも検討してください。

