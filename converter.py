import tkinter as tk
from tkinter import filedialog, messagebox
import os
from pathlib import Path
import tempfile


class MarkdownConverterGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel/Word → Markdown 変換ツール")
        self.root.geometry("600x250")
        self.root.resizable(False, False)
        
        self.selected_file = None
        
        # UI要素の作成
        self._create_widgets()
    
    def _create_widgets(self):
        # ファイルパス表示フレーム
        path_frame = tk.Frame(self.root, pady=20, padx=20)
        path_frame.pack(fill=tk.X)
        
        tk.Label(path_frame, text="選択されたファイル:", font=("Arial", 10)).pack(anchor=tk.W)
        
        self.file_path_var = tk.StringVar(value="ファイルが選択されていません")
        self.file_path_label = tk.Label(
            path_frame, 
            textvariable=self.file_path_var, 
            font=("Arial", 9),
            fg="gray",
            wraplength=550,
            justify=tk.LEFT
        )
        self.file_path_label.pack(anchor=tk.W, pady=5)
        
        # ボタンフレーム
        button_frame = tk.Frame(self.root, pady=10)
        button_frame.pack()
        
        # 参照ボタン
        self.browse_button = tk.Button(
            button_frame,
            text="📁 参照",
            command=self.browse_file,
            width=15,
            height=2,
            font=("Arial", 11, "bold"),
            bg="#4CAF50",
            fg="white",
            cursor="hand2"
        )
        self.browse_button.pack(side=tk.LEFT, padx=10)
        
        # 変換ボタン
        self.convert_button = tk.Button(
            button_frame,
            text="🔄 変換",
            command=self.convert_file,
            width=15,
            height=2,
            font=("Arial", 11, "bold"),
            bg="#2196F3",
            fg="white",
            cursor="hand2",
            state=tk.DISABLED
        )
        self.convert_button.pack(side=tk.LEFT, padx=10)
        
        # ステータス表示
        status_frame = tk.Frame(self.root, pady=10)
        status_frame.pack()
        
        self.status_var = tk.StringVar(value="")
        self.status_label = tk.Label(
            status_frame,
            textvariable=self.status_var,
            font=("Arial", 9),
            fg="blue"
        )
        self.status_label.pack()
    
    def browse_file(self):
        """ファイル選択ダイアログを表示"""
        filetypes = [
            ("Excel/Word ファイル", "*.xlsx *.xls *.docx *.doc"),
            ("Excel ファイル", "*.xlsx *.xls"),
            ("Word ファイル", "*.docx *.doc"),
            ("すべてのファイル", "*.*")
        ]
        
        filename = filedialog.askopenfilename(
            title="変換するファイルを選択",
            filetypes=filetypes
        )
        
        if filename:
            self.selected_file = filename
            self.file_path_var.set(filename)
            self.file_path_label.config(fg="black")
            self.convert_button.config(state=tk.NORMAL)
            self.status_var.set("")
    
    def get_file_type(self, filename):
        """ファイルの拡張子から種類を判定"""
        ext = Path(filename).suffix.lower()
        if ext in ['.xlsx', '.xls']:
            return 'excel'
        elif ext in ['.docx', '.doc']:
            return 'word'
        else:
            return None
    
    def convert_to_pdf(self, input_file, output_pdf):
        """Excel/WordファイルをPDFに変換"""
        file_type = self.get_file_type(input_file)
        
        if file_type == 'excel':
            # pywin32を使用してExcelをPDFに変換
            return self._excel_to_pdf(input_file, output_pdf)
        elif file_type == 'word':
            # pywin32を使用してWordをPDFに変換
            return self._word_to_pdf(input_file, output_pdf)
        else:
            raise ValueError(f"未対応のファイル形式です: {Path(input_file).suffix}")
    
    def _excel_to_pdf(self, input_file, output_pdf):
        """pywin32を使用してExcelをPDFに変換"""
        try:
            import win32com.client
            import pythoncom
            
            # COMの初期化
            pythoncom.CoInitialize()
            
            try:
                # 絶対パスに変換
                input_file = os.path.abspath(input_file)
                output_pdf = os.path.abspath(output_pdf)
                
                # Excelアプリケーションを起動
                excel = win32com.client.Dispatch("Excel.Application")
                excel.Visible = False
                excel.DisplayAlerts = False
                
                try:
                    # Excelファイルを開く
                    workbook = excel.Workbooks.Open(input_file)
                    
                    # PDFとして保存
                    # 0 = xlTypePDF
                    workbook.ExportAsFixedFormat(0, output_pdf)
                    
                    # ワークブックを閉じる
                    workbook.Close(False)
                    
                finally:
                    # Excelアプリケーションを終了
                    excel.Quit()
                    
            finally:
                # COMの終了処理
                pythoncom.CoUninitialize()
            
            return True
            
        except ImportError:
            raise Exception("pywin32がインストールされていません。'pip install pywin32'を実行してください")
        except Exception as e:
            raise Exception(f"Excel PDF変換エラー: {str(e)}")
    
    def _word_to_pdf(self, input_file, output_pdf):
        """pywin32を使用してWordをPDFに変換"""
        try:
            import win32com.client
            import pythoncom
            
            # COMの初期化
            pythoncom.CoInitialize()
            
            try:
                # 絶対パスに変換
                input_file = os.path.abspath(input_file)
                output_pdf = os.path.abspath(output_pdf)
                
                # Wordアプリケーションを起動
                word = win32com.client.Dispatch("Word.Application")
                word.Visible = False
                
                try:
                    # Wordファイルを開く
                    doc = word.Documents.Open(input_file)
                    
                    # PDFとして保存
                    # 17 = wdFormatPDF
                    doc.SaveAs(output_pdf, FileFormat=17)
                    
                    # ドキュメントを閉じる
                    doc.Close(False)
                    
                finally:
                    # Wordアプリケーションを終了
                    word.Quit()
                    
            finally:
                # COMの終了処理
                pythoncom.CoUninitialize()
            
            return True
            
        except ImportError:
            raise Exception("pywin32がインストールされていません。'pip install pywin32'を実行してください")
        except Exception as e:
            raise Exception(f"Word PDF変換エラー: {str(e)}")
    
    def pdf_to_markdown(self, pdf_file, output_md):
        """PDFをMarkdownに変換"""
        try:
            # pymupdfを使用してPDFからテキストを抽出し、Markdownに変換
            import fitz  # PyMuPDF
            
            doc = fitz.open(pdf_file)
            markdown_content = []
            
            markdown_content.append(f"# {Path(pdf_file).stem}\n\n")
            
            for page_num in range(len(doc)):
                page = doc[page_num]
                text = page.get_text()
                
                if text.strip():
                    markdown_content.append(f"## ページ {page_num + 1}\n\n")
                    markdown_content.append(text)
                    markdown_content.append("\n\n---\n\n")
            
            doc.close()
            
            # Markdownファイルに書き込み
            with open(output_md, 'w', encoding='utf-8') as f:
                f.write(''.join(markdown_content))
            
            return True
            
        except ImportError:
            raise Exception("PyMuPDFがインストールされていません")
        except Exception as e:
            raise Exception(f"Markdown変換エラー: {str(e)}")
    
    def convert_file(self):
        """選択されたファイルをMarkdownに変換"""
        if not self.selected_file:
            messagebox.showwarning("警告", "ファイルが選択されていません")
            return
        
        # ファイルの存在確認
        if not os.path.exists(self.selected_file):
            messagebox.showerror("エラー", "選択されたファイルが見つかりません")
            return
        
        # ファイル形式の確認
        file_type = self.get_file_type(self.selected_file)
        if not file_type:
            messagebox.showerror("エラー", "未対応のファイル形式です")
            return
        
        try:
            self.status_var.set("変換中...")
            self.convert_button.config(state=tk.DISABLED)
            self.root.update()
            
            # 出力ファイル名を決定
            input_path = Path(self.selected_file)
            output_md = input_path.parent / f"{input_path.stem}.md"
            
            # 一時PDFファイルを作成
            with tempfile.NamedTemporaryFile(suffix='.pdf', delete=False) as tmp_pdf:
                temp_pdf_path = tmp_pdf.name
            
            try:
                # ステップ1: Excel/Word → PDF
                self.status_var.set("PDFに変換中...")
                self.root.update()
                self.convert_to_pdf(self.selected_file, temp_pdf_path)
                
                # ステップ2: PDF → Markdown
                self.status_var.set("Markdownに変換中...")
                self.root.update()
                self.pdf_to_markdown(temp_pdf_path, str(output_md))
                
                # 一時PDFファイルを削除
                os.unlink(temp_pdf_path)
                
                self.status_var.set(f"✓ 変換完了: {output_md.name}")
                messagebox.showinfo(
                    "成功",
                    f"変換が完了しました!\n\n出力先:\n{output_md}"
                )
                
            finally:
                # 一時ファイルのクリーンアップ
                if os.path.exists(temp_pdf_path):
                    try:
                        os.unlink(temp_pdf_path)
                    except:
                        pass
            
        except Exception as e:
            self.status_var.set("✗ 変換失敗")
            messagebox.showerror("エラー", f"変換に失敗しました:\n{str(e)}")
        
        finally:
            self.convert_button.config(state=tk.NORMAL)


def main():
    root = tk.Tk()
    app = MarkdownConverterGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()

