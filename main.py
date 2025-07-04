import tkinter as tk
from tkinter import filedialog
import os
from openpyxl import load_workbook

class BomExcelSortingApp:
    def __init__(self, root):
        self.root = root
        self.root.title('BOM Excel Sorting')
        
        # 資料夾選擇
        self.folder_path = tk.StringVar()
        tk.Label(root, text='資料夾路徑:').grid(row=0, column=0, padx=5, pady=5, sticky='e')
        tk.Entry(root, textvariable=self.folder_path, width=40, state='readonly').grid(row=0, column=1, padx=5, pady=5)
        tk.Button(root, text='選擇', command=self.select_folder).grid(row=0, column=2, padx=5, pady=5)

        # 主料號
        tk.Label(root, text='主料號:').grid(row=1, column=0, padx=5, pady=5, sticky='e')
        self.main_part = tk.Entry(root, width=40)
        self.main_part.grid(row=1, column=1, padx=5, pady=5, columnspan=2)

        # 替代料
        tk.Label(root, text='替代料:').grid(row=2, column=0, padx=5, pady=5, sticky='e')
        self.alt_part = tk.Entry(root, width=40)
        self.alt_part.grid(row=2, column=1, padx=5, pady=5, columnspan=2)

        # 確認按鈕
        tk.Button(root, text='確認', command=self.confirm).grid(row=3, column=0, columnspan=3, pady=10)

    def select_folder(self):
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            self.folder_path.set(folder_selected)

    def process_excels(self, folder, main_part, alt_part):
        for root, dirs, files in os.walk(folder):
            for file in files:
                if file.lower().endswith('.xlsx') and not file.startswith('~$'):
                    file_path = os.path.join(root, file)
                    print(f'找到 Excel 檔案: {file_path}')  # debug print
                    try:
                        wb = load_workbook(file_path)
                        modified = False
                        for ws in wb.worksheets:
                            for row in ws.iter_rows():
                                for cell in row:
                                    if cell.value == main_part:
                                        ws.insert_rows(cell.row + 1)
                                        ws.cell(row=cell.row + 1, column=cell.column, value=alt_part)
                                        modified = True
                        if modified:
                            wb.save(file_path)
                            print(f'已修改: {file_path}')
                    except Exception as e:
                        print(f'處理 {file_path} 時發生錯誤: {e}')

    def confirm(self):
        folder = self.folder_path.get()
        main_part = self.main_part.get()
        alt_part = self.alt_part.get()
        if not folder or not main_part or not alt_part:
            print('請填寫所有欄位')
            return
        self.process_excels(folder, main_part, alt_part)
        print('處理完成')

if __name__ == '__main__':
    root = tk.Tk()
    app = BomExcelSortingApp(root)
    root.mainloop()
