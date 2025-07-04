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

## 打包成 EXE（含版本資訊）

1. 安裝 pyinstaller：
   ```bash
   pip install pyinstaller
   ```
2. 確認專案根目錄有 `version.txt`（已提供範例，可自訂版本號與產品資訊）
3. 在專案資料夾執行下列指令：
   ```bash
   pyinstaller --onefile --noconsole --version-file=version.txt main.py
   ```
   - 產生的 exe 檔會在 dist 資料夾內。
   - exe 右鍵屬性可看到你設定的版本資訊。

version.txt 範例內容：
```text
# UTF-8
VSVersionInfo(
  ffi=FixedFileInfo(
    filevers=(1,0,0,0),
    prodvers=(1,0,0,0),
    mask=0x3f,
    flags=0x0,
    OS=0x4,
    fileType=0x1,
    subtype=0x0,
    date=(0, 0)
    ),
  kids=[
    StringFileInfo([
      StringTable(
        '040904B0',
        [StringStruct('CompanyName', 'Your Company'),
        StringStruct('FileDescription', 'BOM Excel Sorting'),
        StringStruct('FileVersion', '1.0.0.0'),
        StringStruct('InternalName', 'main'),
        StringStruct('OriginalFilename', 'main.exe'),
        StringStruct('ProductName', 'BOM Excel Sorting'),
        StringStruct('ProductVersion', '1.0.0.0')])
      ]),
    VarFileInfo([VarStruct('Translation', [1033, 1200])])
  ]
)
```
