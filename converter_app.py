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
        self.root.geometry("600x250")
        self.root.resizable(False, False)
        
        self.selected_file = None
        
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
            font=("Arial", 16, "bold")
        )
        title_label.grid(row=0, column=0, columnspan=2, pady=(0, 20))
        
        # ファイル選択セクション
        file_frame = ttk.LabelFrame(main_frame, text="ファイル選択", padding="10")
        file_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 20))
        
        self.file_label = ttk.Label(
            file_frame, 
            text="ファイルが選択されていません",
            wraplength=500
        )
        self.file_label.grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
        
        # 参照ボタン
        browse_button = ttk.Button(
            file_frame,
            text="参照",
            command=self.browse_file,
            width=15
        )
        browse_button.grid(row=0, column=1)
        
        # 変換ボタン
        convert_button = ttk.Button(
            main_frame,
            text="変換",
            command=self.convert_file,
            width=20
        )
        convert_button.grid(row=2, column=0, columnspan=2, pady=(0, 10))
        
        # ステータスバー
        self.status_label = ttk.Label(
            main_frame,
            text="ファイルを選択して変換ボタンを押してください",
            relief=tk.SUNKEN,
            anchor=tk.W
        )
        self.status_label.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E))
        
    def browse_file(self):
        """ファイル参照ダイアログを開く"""
        file_path = filedialog.askopenfilename(
            title="変換するファイルを選択",
            filetypes=[
                ("Office Files", "*.xlsx;*.xls;*.docx;*.doc"),
                ("Excel Files", "*.xlsx;*.xls"),
                ("Word Files", "*.docx;*.doc"),
                ("All Files", "*.*")
            ]
        )
        
        if file_path:
            self.selected_file = file_path
            self.file_label.config(text=f"選択ファイル: {os.path.basename(file_path)}")
            self.status_label.config(text=f"ファイルが選択されました: {file_path}")
    
    def convert_file(self):
        """ファイルをMarkdownに変換"""
        if not self.selected_file:
            messagebox.showwarning("警告", "ファイルが選択されていません")
            return
        
        if not os.path.exists(self.selected_file):
            messagebox.showerror("エラー", "選択されたファイルが存在しません")
            return
        
        try:
            self.status_label.config(text="変換中...")
            self.root.update()
            
            # 拡張子を取得
            file_ext = Path(self.selected_file).suffix.lower()
            
            # Excel/Word → PDF変換
            pdf_path = self.convert_to_pdf(self.selected_file, file_ext)
            
            # PDF → Markdown変換
            markdown_path = self.convert_pdf_to_markdown(pdf_path)
            
            # 一時PDFファイルを削除
            if os.path.exists(pdf_path):
                os.remove(pdf_path)
            
            self.status_label.config(text=f"変換完了: {markdown_path}")
            messagebox.showinfo(
                "完了", 
                f"変換が完了しました\n\n出力ファイル:\n{markdown_path}"
            )
            
        except Exception as e:
            self.status_label.config(text="エラーが発生しました")
            messagebox.showerror("エラー", f"変換中にエラーが発生しました:\n{str(e)}")
    
    def convert_to_pdf(self, file_path, file_ext):
        """Excel/WordファイルをPDFに変換"""
        # COMの初期化
        pythoncom.CoInitialize()
        
        try:
            output_pdf = str(Path(file_path).with_suffix('.pdf'))
            abs_file_path = os.path.abspath(file_path)
            abs_output_pdf = os.path.abspath(output_pdf)
            
            if file_ext in ['.xlsx', '.xls']:
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
                    
            elif file_ext in ['.docx', '.doc']:
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
        output_md = str(Path(pdf_path).with_suffix('.md'))
        
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
            for img_index, img in enumerate(image_list):
                xref = img[0]
                base_image = doc.extract_image(xref)
                image_bytes = base_image["image"]
                image_ext = base_image["ext"]
                
                # 画像を保存
                image_filename = f"{Path(pdf_path).stem}_page{page_num}_img{img_index + 1}.{image_ext}"
                image_path = Path(pdf_path).parent / image_filename
                
                with open(image_path, "wb") as img_file:
                    img_file.write(image_bytes)
                
                # Markdownに画像参照を追加
                markdown_content.append(f"![Image]({image_filename})\n\n")
        
        doc.close()
        
        # Markdownファイルを保存
        with open(output_md, 'w', encoding='utf-8') as f:
            f.writelines(markdown_content)
        
        return output_md


def main():
    """メイン関数"""
    root = tk.Tk()
    app = FileConverterApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()

