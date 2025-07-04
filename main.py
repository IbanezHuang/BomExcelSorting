import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
from openpyxl import load_workbook
import datetime
from openpyxl.styles import Font

class BomExcelSortingApp:
    def __init__(self, root):
        self.root = root
        self.root.title('BOM Excel Sorting')
        
        # 設定主題顏色與字型
        self.root.configure(bg='#f0f4f8')
        label_font = ('Microsoft JhengHei', 12)
        entry_font = ('Microsoft JhengHei', 11)
        btn_font = ('Microsoft JhengHei', 11, 'bold')
        style = ttk.Style()
        style.theme_use('clam')
        style.configure('TButton', font=btn_font, padding=6)
        style.configure('TLabel', font=label_font, background='#f0f4f8')
        style.configure('TEntry', font=entry_font)
        style.configure('TProgressbar', thickness=18, troughcolor='#e0e7ef', background='#4a90e2')

        # 資料夾選擇
        self.folder_path = tk.StringVar()  # <--- 修正：先宣告 self.folder_path
        ttk.Label(root, text='資料夾路徑:').grid(row=0, column=0, padx=8, pady=8, sticky='e')
        ttk.Entry(root, textvariable=self.folder_path, width=40, state='readonly').grid(row=0, column=1, padx=8, pady=8)
        ttk.Button(root, text='選擇', command=self.select_folder).grid(row=0, column=2, padx=8, pady=8)

        # 另存資料夾選擇
        self.save_folder_path = tk.StringVar()
        ttk.Label(root, text='另存資料夾路徑:').grid(row=1, column=0, padx=8, pady=8, sticky='e')
        ttk.Entry(root, textvariable=self.save_folder_path, width=40, state='readonly').grid(row=1, column=1, padx=8, pady=8)
        ttk.Button(root, text='選擇', command=self.select_save_folder).grid(row=1, column=2, padx=8, pady=8)

        # 主料號
        ttk.Label(root, text='主料號:').grid(row=2, column=0, padx=8, pady=8, sticky='e')
        self.main_part = ttk.Entry(root, width=40)
        self.main_part.grid(row=2, column=1, padx=8, pady=8, columnspan=2)

        # 替代料
        ttk.Label(root, text='替代料:').grid(row=3, column=0, padx=8, pady=8, sticky='e')
        self.alt_part = ttk.Entry(root, width=40)
        self.alt_part.grid(row=3, column=1, padx=8, pady=8, columnspan=2)

        # 替代料品名
        ttk.Label(root, text='替代料品名:').grid(row=4, column=0, padx=8, pady=8, sticky='e')
        self.alt_part_name = ttk.Entry(root, width=40)
        self.alt_part_name.grid(row=4, column=1, padx=8, pady=8, columnspan=2)

        # 替代料規格
        ttk.Label(root, text='替代料規格:').grid(row=5, column=0, padx=8, pady=8, sticky='e')
        self.alt_part_spec = ttk.Entry(root, width=40)
        self.alt_part_spec.grid(row=5, column=1, padx=8, pady=8, columnspan=2)

        # 進度條
        self.progress = ttk.Progressbar(root, orient="horizontal", length=340, mode="determinate", style='TProgressbar')
        self.progress.grid(row=7, column=0, columnspan=3, pady=16)
        self.progress.grid_remove()

        # 確認、取消、關閉按鈕
        ttk.Button(root, text='確認', command=self.confirm).grid(row=6, column=0, pady=12, padx=4, sticky='ew')
        ttk.Button(root, text='取消', command=self.cancel).grid(row=6, column=1, pady=12, padx=4, sticky='ew')
        ttk.Button(root, text='關閉', command=self.close_app).grid(row=6, column=2, pady=12, padx=4, sticky='ew')

    def select_folder(self):
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            self.folder_path.set(folder_selected)

    def select_save_folder(self):
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            self.save_folder_path.set(folder_selected)

    def process_excels(self, folder, main_part, alt_part):
        alt_part_name = self.alt_part_name.get()
        alt_part_spec = self.alt_part_spec.get()
        save_folder = self.save_folder_path.get()
        # 收集所有 Excel 檔案
        excel_files = []
        for root_dir, dirs, files in os.walk(folder):
            for file in files:
                if file.lower().endswith('.xlsx') and not file.startswith('~$'):
                    excel_files.append(os.path.join(root_dir, file))
        total = len(excel_files)
        if total == 0:
            return
        self.progress['maximum'] = total
        self.progress['value'] = 0
        self.progress.grid()
        log_path = os.path.join(os.getcwd(), "process_log.txt")
        with open(log_path, 'a', encoding='utf-8') as log_file:
            for idx, file_path in enumerate(excel_files, 1):
                try:
                    wb = load_workbook(file_path)
                    modified = False
                    for ws in wb.worksheets:
                        max_row = ws.max_row
                        max_col = ws.max_column
                        for row in ws.iter_rows():
                            for cell in row:
                                if cell.value == main_part:
                                    # 先檢查下方到下一個主料號前有無相同替代料
                                    insert_row = cell.row + 1
                                    found_duplicate = False
                                    check_row = insert_row
                                    while check_row <= max_row:
                                        a1_value = ws.cell(row=check_row, column=1).value
                                        alt_value = ws.cell(row=check_row, column=cell.column).value
                                        # 遇到下一個主料號（A欄有值）就停止
                                        if a1_value:
                                            break
                                        # 檢查同欄有無相同替代料
                                        if alt_value == alt_part:
                                            found_duplicate = True
                                            break
                                        check_row += 1
                                    if not found_duplicate:
                                        ws.insert_rows(insert_row)
                                        ws.cell(row=insert_row, column=cell.column, value=alt_part)
                                        ws.cell(row=insert_row, column=cell.column + 1, value=alt_part_name)
                                        ws.cell(row=insert_row, column=cell.column + 2, value=alt_part_spec)
                                        ws.cell(row=insert_row, column=cell.column + 11, value=datetime.datetime.now().strftime('%Y-%m-%d') + ' 客戶確認可替代')
                                        # 設定字體為細明體，大小8，紅色
                                        # 設定字體為細明體，大小8
                                        font_normal = Font(name='MingLiU', size=8)
                                        font_red = Font(name='MingLiU', size=8, color="FFFF0000", bold=True)
                                        ws.cell(row=insert_row, column=cell.column).font = font_normal
                                        ws.cell(row=insert_row, column=cell.column + 1).font = font_normal
                                        ws.cell(row=insert_row, column=cell.column + 2).font = font_normal
                                        ws.cell(row=insert_row, column=cell.column + 11).font = font_red
                                        modified = True
                    if modified:
                        if save_folder:
                            base = os.path.basename(file_path)
                            name, ext = os.path.splitext(base)
                            date_str = datetime.datetime.now().strftime('%Y%m%d')
                            new_name = f"{name}-{date_str}{ext}"
                            save_path = os.path.join(save_folder, new_name)
                            wb.save(save_path)
                        else:
                            wb.save(file_path)
                    # 寫入 log
                    log_file.write(f"{datetime.datetime.now()} - {file_path}\n")
                except Exception as e:
                    log_file.write(f"{datetime.datetime.now()} - {file_path} - Error: {e}\n")
                self.progress['value'] = idx
                self.root.update_idletasks()
        self.progress.grid_remove()
        messagebox.showinfo("完成", "所有檔案處理完成！")

    def confirm(self):
        folder = self.folder_path.get()
        main_part = self.main_part.get()
        alt_part = self.alt_part.get()
        if not folder or not main_part or not alt_part:
            messagebox.showwarning('警告', '請填寫所有欄位')
            return
        self.process_excels(folder, main_part, alt_part)

    def cancel(self):
        # 清空所有欄位
        self.folder_path.set("")
        self.main_part.delete(0, tk.END)
        self.alt_part.delete(0, tk.END)
        self.alt_part_name.delete(0, tk.END)
        self.alt_part_spec.delete(0, tk.END)
        self.progress.grid_remove()

    def close_app(self):
        self.root.destroy()

if __name__ == '__main__':
    root = tk.Tk()
    app = BomExcelSortingApp(root)
    root.mainloop()
