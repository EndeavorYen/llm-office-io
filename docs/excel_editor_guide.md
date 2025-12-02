# Excel 編輯器使用指南

## 📖 工具介紹

`excel_editor.py` 是一個強大的 Excel 互動式編輯工具，讓您可以透過簡單的命令列指令修改 Excel 檔案。

## 🎯 主要功能

### 1. 列出工作表
顯示所有工作表名稱和基本資訊

```bash
python src/excel_editor.py data.xlsx list
```

### 2. 查看工作表內容
顯示指定工作表的資料

```bash
# 查看活動工作表（前 10 行）
python src/excel_editor.py data.xlsx view

# 查看指定工作表
python src/excel_editor.py data.xlsx view Sheet1

# 指定顯示行數
python src/excel_editor.py data.xlsx view Sheet1 --max-rows 20
```

### 3. 替換文字
在工作表中搜尋並替換文字

```bash
# 替換所有工作表的文字
python src/excel_editor.py data.xlsx replace "舊值" "新值"

# 只替換特定工作表
python src/excel_editor.py data.xlsx replace "舊值" "新值" --sheet Sheet1
```

### 4. 更新儲存格
修改指定儲存格的值

```bash
python src/excel_editor.py data.xlsx update-cell Sheet1 A1 "新值"
python src/excel_editor.py data.xlsx update-cell Sheet1 B2 "100"
```

### 5. 新增行
在工作表中插入新行

```bash
# 在最後新增行
python src/excel_editor.py data.xlsx add-row Sheet1 "數據1" "數據2" "數據3"

# 在指定位置插入
python src/excel_editor.py data.xlsx add-row Sheet1 "數據1" "數據2" --position 2
```

### 6. 刪除行
刪除指定行

```bash
python src/excel_editor.py data.xlsx delete-row Sheet1 5
```

### 7. 搜尋儲存格
搜尋包含特定文字的儲存格

```bash
# 搜尋所有工作表
python src/excel_editor.py data.xlsx find "關鍵字"

# 只搜尋特定工作表
python src/excel_editor.py data.xlsx find "關鍵字" --sheet Sheet1
```

## 💾 輸出設定

預設會覆蓋原檔案，如果要另存新檔：

```bash
python src/excel_editor.py input.xlsx replace "A" "B" --output output.xlsx
```

## 📝 使用範例

### 範例 1：批次更新價格

```bash
# 1. 查看價格表
python src/excel_editor.py products.xlsx view PriceList

# 2. 將所有價格從 $100 改為 $120
python src/excel_editor.py products.xlsx replace "$100" "$120" --sheet PriceList

# 3. 更新標題
python src/excel_editor.py products.xlsx update-cell PriceList A1 "2025 年價格表"
```

### 範例 2：資料維護

```bash
# 1. 搜尋待處理項目
python src/excel_editor.py tasks.xlsx find "待處理"

# 2. 新增新任務
python src/excel_editor.py tasks.xlsx add-row Tasks "新任務" "高優先級" "2025-12-10"

# 3. 刪除完成的任務（假設在第 3 行）
python src/excel_editor.py tasks.xlsx delete-row Tasks 3
```

### 範例 3：報表更新

```bash
# 1. 列出所有工作表
python src/excel_editor.py report.xlsx list

# 2. 更新統計數據
python src/excel_editor.py report.xlsx update-cell Summary B5 "95%"

# 3. 替換報告日期
python src/excel_editor.py report.xlsx replace "2024-12-01" "2025-12-02"
```

## ⚠️ 注意事項

1. **備份重要文件** - 修改前建議先備份
2. **支援格式** - 只支援 `.xlsx` 格式（Excel 2007+）
3. **儲存格參照** - 使用標準格式如 A1, B2, C3
4. **不支援功能** - 不支援編輯公式、圖表、巨集

## 🔧 技術細節

### 依賴套件

- **openpyxl** - Excel 檔案操作

### 系統需求

- Python 3.8+
- Windows / Linux / macOS

---

**最後更新**: 2025-12-02
