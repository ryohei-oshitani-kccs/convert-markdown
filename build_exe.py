"""
Excel/Word → Markdown 変換ツールのexeビルドスクリプト

このスクリプトはWindows環境で実行してください。
PyInstallerを使用して実行ファイル(.exe)を作成します。
"""

import PyInstaller.__main__
import os
import sys

def build_exe():
    """PyInstallerを使用してexeファイルをビルド"""
    
    print("=" * 60)
    print("Excel/Word → Markdown 変換ツール exe ビルドスクリプト")
    print("=" * 60)
    print()
    
    # ビルド設定
    app_name = "MarkdownConverter"
    main_script = "converter.py"
    icon_file = None  # アイコンファイルがあればここで指定
    
    # 確認
    if not os.path.exists(main_script):
        print(f"エラー: {main_script} が見つかりません")
        sys.exit(1)
    
    print(f"ビルド対象: {main_script}")
    print(f"出力名: {app_name}.exe")
    print()
    print("ビルドを開始します...")
    print()
    
    # PyInstallerの引数を構築
    pyinstaller_args = [
        main_script,
        '--name', app_name,
        '--onefile',                    # 単一の実行ファイルとして作成
        '--windowed',                   # コンソールウィンドウを表示しない
        '--noconfirm',                  # 既存のファイルを上書き
        '--clean',                      # ビルド前にキャッシュをクリア
        
        # 必要なモジュールを明示的に含める
        '--hidden-import', 'win32com.client',
        '--hidden-import', 'pythoncom',
        '--hidden-import', 'pywintypes',
        '--hidden-import', 'fitz',
        '--hidden-import', 'pymupdf',
        
        # tkinterの必要なモジュール
        '--hidden-import', 'tkinter',
        '--hidden-import', 'tkinter.filedialog',
        '--hidden-import', 'tkinter.messagebox',
        
        # 追加データファイル（必要に応じて）
        # '--add-data', 'data;data',
        
        # デバッグ情報を含める（問題がある場合はコメントアウト）
        # '--debug', 'all',
    ]
    
    # アイコンファイルがある場合は追加
    if icon_file and os.path.exists(icon_file):
        pyinstaller_args.extend(['--icon', icon_file])
    
    # ビルド実行
    try:
        PyInstaller.__main__.run(pyinstaller_args)
        print()
        print("=" * 60)
        print("✓ ビルドが完了しました！")
        print("=" * 60)
        print()
        print(f"実行ファイルの場所: dist/{app_name}.exe")
        print()
        print("このexeファイルをWindows環境で実行できます。")
        print("注意: Microsoft Officeがインストールされている必要があります。")
        print()
        
    except Exception as e:
        print()
        print("=" * 60)
        print("✗ ビルドに失敗しました")
        print("=" * 60)
        print(f"エラー: {str(e)}")
        sys.exit(1)


if __name__ == "__main__":
    build_exe()

