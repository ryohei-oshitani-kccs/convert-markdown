"""
Excel/Word to Markdown Converter
Excel/WordファイルをPDF経由でMarkdownに変換するGUIアプリケーション
"""

import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pathlib import Path
import win32com.client
import fitz  # PyMuPDF
import pythoncom


class FileConverterApp:
    """ファイル変換アプリケーションのメインクラス"""

    def __init__(self, root):
        """アプリケーションの初期化"""
        self.root = root
        self.root.title("Excel/Word to Markdown Converter")
        self.root.geometry("700x320")
        self.root.resizable(False, False)

        self.selected_files = []  # 複数ファイルに対応

        self.setup_ui()

    def setup_ui(self):
        """UIコンポーネントのセットアップ"""
        # メインフレーム
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # タイトル
        title_label = ttk.Label(
            main_frame,
            text="Excel/Word to Markdown Converter",
            font=("Arial", 16, "bold"),
        )
        title_label.grid(row=0, column=0, columnspan=2, pady=(0, 20))

        # ファイル選択セクション
        file_frame = ttk.LabelFrame(main_frame, text="ファイル選択", padding="10")
        file_frame.grid(
            row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 20)
        )

        self.file_label = ttk.Label(
            file_frame, text="ファイルが選択されていません", wraplength=500
        )
        self.file_label.grid(row=0, column=0, columnspan=3, sticky=tk.W, pady=(0, 10))

        # 1ファイル参照ボタン
        browse_button = ttk.Button(
            file_frame, text="1ファイル選択", command=self.browse_file, width=18
        )
        browse_button.grid(row=1, column=0, padx=(0, 5))

        # 複数ファイル参照ボタン
        browse_multiple_button = ttk.Button(
            file_frame,
            text="複数ファイル選択",
            command=self.browse_multiple_files,
            width=18,
        )
        browse_multiple_button.grid(row=1, column=1, padx=(0, 5))

        # フォルダ選択ボタン
        browse_folder_button = ttk.Button(
            file_frame, text="フォルダ選択", command=self.browse_folder, width=18
        )
        browse_folder_button.grid(row=1, column=2)

        # 変換ボタン
        convert_button = ttk.Button(
            main_frame, text="変換", command=self.convert_file, width=20
        )
        convert_button.grid(row=2, column=0, columnspan=2, pady=(0, 10))

        # ステータスバー
        self.status_label = ttk.Label(
            main_frame,
            text="ファイルを選択して変換ボタンを押してください",
            relief=tk.SUNKEN,
            anchor=tk.W,
        )
        self.status_label.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E))

    def browse_file(self):
        """1ファイル参照ダイアログを開く"""
        file_path = filedialog.askopenfilename(
            title="変換するファイルを選択",
            filetypes=[
                ("Office Files", "*.xlsx;*.xls;*.docx;*.doc"),
                ("Excel Files", "*.xlsx;*.xls"),
                ("Word Files", "*.docx;*.doc"),
                ("All Files", "*.*"),
            ],
        )

        if file_path:
            self.selected_files = [file_path]
            self.file_label.config(text=f"選択ファイル: {os.path.basename(file_path)}")
            self.status_label.config(text=f"ファイルが選択されました: {file_path}")

    def browse_multiple_files(self):
        """複数ファイル参照ダイアログを開く"""
        file_paths = filedialog.askopenfilenames(
            title="変換するファイルを複数選択",
            filetypes=[
                ("Office Files", "*.xlsx;*.xls;*.docx;*.doc"),
                ("Excel Files", "*.xlsx;*.xls"),
                ("Word Files", "*.docx;*.doc"),
                ("All Files", "*.*"),
            ],
        )

        if file_paths:
            self.selected_files = list(file_paths)
            file_count = len(self.selected_files)
            if file_count == 1:
                self.file_label.config(
                    text=f"選択ファイル: {os.path.basename(self.selected_files[0])}"
                )
            else:
                self.file_label.config(text=f"{file_count}個のファイルが選択されました")
            self.status_label.config(text=f"{file_count}個のファイルが選択されました")

    def browse_folder(self):
        """フォルダ参照ダイアログを開く"""
        folder_path = filedialog.askdirectory(
            title="変換するファイルが含まれるフォルダを選択"
        )

        if folder_path:
            # フォルダ内のOfficeファイルを検索
            folder = Path(folder_path)
            office_extensions = [".xlsx", ".xls", ".docx", ".doc"]
            self.selected_files = []

            for ext in office_extensions:
                self.selected_files.extend(folder.glob(f"*{ext}"))

            # Pathオブジェクトを文字列に変換
            self.selected_files = [str(f) for f in self.selected_files]

            file_count = len(self.selected_files)
            if file_count == 0:
                self.file_label.config(
                    text="フォルダ内にOfficeファイルが見つかりませんでした"
                )
                self.status_label.config(
                    text="変換可能なファイルが見つかりませんでした"
                )
                messagebox.showinfo(
                    "情報",
                    "選択されたフォルダ内に変換可能なファイルが見つかりませんでした",
                )
            else:
                self.file_label.config(
                    text=f"フォルダ内の{file_count}個のファイルが選択されました"
                )
                self.status_label.config(
                    text=f"フォルダから{file_count}個のファイルが選択されました"
                )

    def convert_file(self):
        """ファイルをMarkdownに変換"""
        if not self.selected_files:
            messagebox.showwarning("警告", "ファイルが選択されていません")
            return

        # 存在しないファイルをチェック
        valid_files = [f for f in self.selected_files if os.path.exists(f)]
        if not valid_files:
            messagebox.showerror("エラー", "選択されたファイルが存在しません")
            return

        if len(valid_files) < len(self.selected_files):
            missing_count = len(self.selected_files) - len(valid_files)
            messagebox.showwarning(
                "警告",
                f"{missing_count}個のファイルが見つかりませんでした。\n存在するファイルのみ変換を続行します。",
            )

        total_files = len(valid_files)
        success_count = 0
        error_files = []

        try:
            for idx, file_path in enumerate(valid_files, start=1):
                try:
                    # 進捗表示
                    self.status_label.config(
                        text=f"変換中... ({idx}/{total_files}) {os.path.basename(file_path)}"
                    )
                    self.root.update()

                    # 拡張子を取得
                    file_ext = Path(file_path).suffix.lower()

                    # Excel/Word → PDF変換
                    pdf_path = self.convert_to_pdf(file_path, file_ext)

                    # PDF → Markdown変換
                    markdown_path = self.convert_pdf_to_markdown(pdf_path)

                    # # 一時PDFファイルを削除
                    # if os.path.exists(pdf_path):
                    #     os.remove(pdf_path)

                    success_count += 1

                except Exception as e:
                    error_files.append((os.path.basename(file_path), str(e)))

            # 結果メッセージ
            self.status_label.config(
                text=f"変換完了: {success_count}/{total_files} ファイル"
            )

            if error_files:
                error_msg = f"変換完了: {success_count}/{total_files} ファイル\n\n"
                error_msg += "以下のファイルでエラーが発生しました:\n"
                for filename, error in error_files:
                    error_msg += f"\n・{filename}\n  {error}\n"
                messagebox.showwarning("完了（エラーあり）", error_msg)
            else:
                messagebox.showinfo(
                    "完了",
                    f"すべてのファイルの変換が完了しました\n\n変換ファイル数: {success_count}",
                )

        except Exception as e:
            self.status_label.config(text="エラーが発生しました")
            messagebox.showerror("エラー", f"変換中にエラーが発生しました:\n{str(e)}")

    def convert_to_pdf(self, file_path, file_ext):
        """Excel/WordファイルをPDFに変換"""
        # COMの初期化
        pythoncom.CoInitialize()

        try:
            output_pdf = str(Path(file_path).with_suffix(".pdf"))
            abs_file_path = os.path.abspath(file_path)
            abs_output_pdf = os.path.abspath(output_pdf)

            if file_ext in [".xlsx", ".xls"]:
                # Excelの場合
                excel = win32com.client.Dispatch("Excel.Application")
                excel.Visible = False
                excel.DisplayAlerts = False

                try:
                    workbook = excel.Workbooks.Open(abs_file_path)
                    workbook.ExportAsFixedFormat(0, abs_output_pdf)
                    workbook.Close(False)
                finally:
                    excel.Quit()

            elif file_ext in [".docx", ".doc"]:
                # Wordの場合
                word = win32com.client.Dispatch("Word.Application")
                word.Visible = False

                try:
                    doc = word.Documents.Open(abs_file_path)
                    doc.SaveAs(abs_output_pdf, FileFormat=17)  # 17 = wdFormatPDF
                    doc.Close(False)
                finally:
                    word.Quit()
            else:
                raise ValueError(f"サポートされていないファイル形式です: {file_ext}")

            return output_pdf

        finally:
            pythoncom.CoUninitialize()

    def convert_pdf_to_markdown(self, pdf_path):
        """PDFファイルをMarkdownに変換"""
        output_md = str(Path(pdf_path).with_suffix(".md"))

        # 画像保存用フォルダの作成
        images_folder_name = f"{Path(pdf_path).stem}_images"
        images_folder_path = Path(pdf_path).parent / images_folder_name

        # PDFを開く
        doc = fitz.open(pdf_path)

        markdown_content = []
        markdown_content.append(f"# {Path(pdf_path).stem}\n\n")

        # 各ページを処理
        for page_num, page in enumerate(doc, start=1):
            # ページタイトル
            if len(doc) > 1:
                markdown_content.append(f"## ページ {page_num}\n\n")

            # テキストを抽出
            text = page.get_text()

            if text.strip():
                markdown_content.append(text)
                markdown_content.append("\n\n")

            # 画像を抽出
            image_list = page.get_images()
            if image_list:
                # 画像がある場合のみフォルダを作成
                images_folder_path.mkdir(exist_ok=True)

            for img_index, img in enumerate(image_list):
                xref = img[0]
                base_image = doc.extract_image(xref)
                image_bytes = base_image["image"]
                image_ext = base_image["ext"]

                # 画像を保存（フォルダ内に保存）
                image_filename = f"{Path(pdf_path).stem}_page{page_num}_img{img_index + 1}.{image_ext}"
                image_path = images_folder_path / image_filename

                with open(image_path, "wb") as img_file:
                    img_file.write(image_bytes)

                # Markdownに画像参照を追加（フォルダ名を含める）
                markdown_content.append(
                    f"![Image]({images_folder_name}/{image_filename})\n\n"
                )

        doc.close()

        # Markdownファイルを保存
        with open(output_md, "w", encoding="utf-8") as f:
            f.writelines(markdown_content)

        return output_md


def main():
    """メイン関数"""
    root = tk.Tk()
    app = FileConverterApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
