# BOM Excel Sorting

本專案提供一個簡易的圖形介面，讓你可以：
- 選擇一個資料夾（會遞迴搜尋所有子資料夾）
- 輸入主料號與替代料
- 自動搜尋所有 Excel 檔案（.xlsx），找到主料號後在其下方插入替代料，並自動存檔

## 使用方式

1. 安裝 Python 3.8 以上
2. 安裝必要套件：
   ```bash
   pip install -r requirements.txt
   ```
3. 執行主程式：
   ```bash
   python main.py
   ```
4. 依照介面操作

## 注意事項
- 只處理 .xlsx 檔案（不處理 .xls）
- 請勿同時開啟 Excel 檔案以避免存檔失敗
- 若遇到權限問題請以管理員身份執行

## 需求套件
- tkinter（Python 內建）
- openpyxl

## 打包成 EXE

1. 安裝 pyinstaller：
   ```bash
   pip install pyinstaller
   ```
2. 在專案資料夾執行下列指令：
   ```bash
   pyinstaller --onefile --noconsole main.py
   ```
   - 產生的 exe 檔會在 dist 資料夾內。
   - 若要自訂圖示，可加參數：
     ```bash
     pyinstaller --onefile --noconsole --icon=icon.ico main.py
     ```

3. 將 dist 資料夾內的 exe 檔案提供給使用者即可。
