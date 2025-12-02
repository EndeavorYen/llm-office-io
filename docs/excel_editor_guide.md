# Excel Editor 完整使用指南

## 📖 簡介

Excel Editor 提供 11 個強大功能，包括工作表管理、格式設定和公式支援。

---

## 🚀 快速開始

```python
from src.excel_editor import ExcelEditor

# 開啟 Excel 檔案
editor = ExcelEditor("data.xlsx")

# 執行操作
editor.replace_text("舊值", "新值")
editor.save("output.xlsx")
```

---

## 📋 功能列表

### 1. 列出工作表 `list_sheets()`

```python
# 顯示所有工作表名稱
editor.list_sheets()
```

---

### 2. 查看工作表內容 `view_sheet()`

```python
# 查看活動工作表（前 10 行）
editor.view_sheet()

# 查看指定工作表
editor.view_sheet("Sheet1")

# 指定顯示行數
editor.view_sheet("Sheet1", max_rows=20)
```

---

### 3. 文字替換 `replace_text()`

```python
# 替換所有工作表的文字
count = editor.replace_text("舊值", "新值")

# 只替換特定工作表
count = editor.replace_text("舊值", "新值", sheet_name="Sheet1")
```

---

### 4. 更新儲存格 `update_cell()`

```python
# 更新指定儲存格
editor.update_cell("Sheet1", "A1", "新值")
editor.update_cell("財務", "B5", 12000)
```

---

### 5. 新增行 `add_row()`

```python
# 在最後新增行
data = ["產品A", 100, 5000]
editor.add_row("Sheet1", data)

# 在指定位置插入
editor.add_row("Sheet1", ["產品B", 200, 8000], position=2)
```

---

### 6. 刪除行 `delete_row()`

```python
# 刪除第 5 行
editor.delete_row("Sheet1", row_number=5)
```

---

### 7. 搜尋儲存格 `find_cells()`

```python
# 搜尋所有工作表
results = editor.find_cells("關鍵字")

# 只搜尋特定工作表
results = editor.find_cells("關鍵字", sheet_name="Sheet1")
```

---

### 8. 新增工作表 `add_sheet()` 🆕

```python
# 在最後新增工作表
editor.add_sheet("新工作表")

# 在特定位置插入
editor.add_sheet("Q1資料", position=0)  # 插入到最前面
```

---

### 9. 刪除工作表 `delete_sheet()` 🆕

```python
# 刪除工作表
editor.delete_sheet("舊工作表")
```

**注意**: 無法刪除唯一的工作表

---

### 10. 設定儲存格格式 `set_cell_format()` 🆕

```python
# 設定粗體、字體大小
editor.set_cell_format(
    sheet_name="Sheet1",
    cell_ref="A1",
    bold=True,
    font_size=14
)

# 設定背景顏色（16進位）
editor.set_cell_format(
    sheet_name="Sheet1",
    cell_ref="B2",
    bg_color="FFFF00",  # 黃色
    alignment="center"
)

# 完整範例
editor.set_cell_format(
    sheet_name="報表",
    cell_ref="C3",
    bold=True,
    font_size=12,
    bg_color="CCE5FF",  # 淺藍色
    alignment="right"
)
```

**常用顏色**:
- 黃色: `"FFFF00"`
- 淺藍: `"CCE5FF"`
- 淺綠: `"CCFFCC"`
- 淺紅: `"FFCCCC"`
- 橙色: `"FFA500"`

**對齊選項**: `'left'`, `'center'`, `'right'`

---

### 11. 設定公式 `set_formula()` 🆕

```python
# SUM 公式
editor.set_formula("Sheet1", "D10", "=SUM(D1:D9)")

# AVERAGE 公式
editor.set_formula("Sheet1", "E10", "=AVERAGE(E1:E9)")

# 其他公式
editor.set_formula("Sheet1", "F5", "=A5*B5")
editor.set_formula("Sheet1", "G1", "=IF(A1>100,\"高\",\"低\")")
```

---

### 12. 儲存檔案 `save()`

```python
# 覆蓋原檔案
editor.save()

# 另存新檔
editor.save("output.xlsx")
```

---

## 💡 實用範例

### 範例 1: 季度報表製作

```python
editor = ExcelEditor("report.xlsx")

# 新增 Q1 工作表
editor.add_sheet("Q1_2025", position=0)

# 設定標題
editor.update_cell("Q1_2025", "A1", "Q1 2025 財務報表")
editor.set_cell_format(
    "Q1_2025", "A1",
    bold=True,
    font_size=16,
    bg_color="4472C4",  # 深藍
    alignment="center"
)

# 添加數據
headers = ["月份", "收入", "支出", "淨利"]
editor.add_row("Q1_2025", headers)

data = [
    ["1月", 100000, 60000, 40000],
    ["2月", 120000, 70000, 50000],
    ["3月", 115000, 65000, 50000]
]

for row in data:
    editor.add_row("Q1_2025", row)

# 設定總計公式
editor.update_cell("Q1_2025", "A6", "總計")
editor.set_formula("Q1_2025", "B6", "=SUM(B3:B5)")
editor.set_formula("Q1_2025", "C6", "=SUM(C3:C5)")
editor.set_formula("Q1_2025", "D6", "=SUM(D3:D5)")

# 格式化總計行
for col in ["A6", "B6", "C6", "D6"]:
    editor.set_cell_format(
        "Q1_2025", col,
        bold=True,
        bg_color="D9E1F2"
    )

editor.save()
```

---

### 範例 2: 批次數據更新

```python
editor = ExcelEditor("products.xlsx")

# 更新所有價格（+10%）
# 先搜尋所有價格儲存格
results = editor.find_cells("$", sheet_name="Price List")

for sheet, cell_ref, value in results:
    if isinstance(value, str) and "$" in value:
        # 提取數字並增加 10%
        old_price = float(value.replace("$", ""))
        new_price = old_price * 1.1
        editor.update_cell(sheet, cell_ref, f"${new_price:.2f}")

# 更新日期
editor.replace_text("2024", "2025", sheet_name="Price List")

# 標記為已更新
editor.update_cell("Price List", "A1", "價格表 (2025年1月更新)")
editor.set_cell_format(
    "Price List", "A1",
    bold=True,
    bg_color="FFFF00"
)

editor.save()
```

---

### 範例 3: 工作表整理

```python
editor = ExcelEditor("data.xlsx")

# 刪除舊工作表
old_sheets = ["2022資料", "2023資料", "暫存"]
for sheet in old_sheets:
    try:
        editor.delete_sheet(sheet)
    except:
        pass

# 新增當年度工作表
for quarter in ["Q1", "Q2", "Q3", "Q4"]:
    sheet_name = f"2025_{quarter}"
    editor.add_sheet(sheet_name)
    
    # 設定標題
    editor.update_cell(sheet_name, "A1", f"2025 年 {quarter} 資料")
    editor.set_cell_format(
        sheet_name, "A1",
        bold=True,
        font_size=14,
        alignment="center"
    )

editor.save()
```

---

### 範例 4: 自動化報表格式

```python
editor = ExcelEditor("monthly_report.xlsx")

# 格式化標題行
headers = ["A1", "B1", "C1", "D1", "E1"]
for cell in headers:
    editor.set_cell_format(
        "Report", cell,
        bold=True,
        font_size=12,
        bg_color="366092",  # 深藍
        alignment="center"
    )

# 格式化數據區域（使用淺色背景）
for row in range(2, 12):  # 行 2-11
    bg = "F2F2F2" if row % 2 == 0 else "FFFFFF"  # 斑馬紋
    for col in ["A", "B", "C", "D", "E"]:
        cell_ref = f"{col}{row}"
        editor.set_cell_format(
            "Report", cell_ref,
            bg_color=bg,
            alignment="left"
        )

# 添加總計行
editor.set_formula("Report", "E12", "=SUM(E2:E11)")
editor.set_cell_format(
    "Report", "E12",
    bold=True,
    bg_color="FFD966"  # 黃色
)

editor.save()
```

---

## ⚠️ 注意事項

1. **檔案格式**: 僅支援 `.xlsx` 格式
2. **儲存格參照**: 使用標準格式（A1, B2, C3...）
3. **行號從 1 開始**: 第一行是 1（不是 0）
4. **工作表名稱**: 不可重複
5. **顏色格式**: 使用 6 位 16 進位（如 FFFF00）

---

## 🎨 常用顏色代碼

| 顏色 | 16進位碼 |
|------|----------|
| 黃色 | FFFF00 |
| 橙色 | FFA500 |
| 紅色 | FF0000 |
| 粉紅 | FFC0CB |
| 綠色 | 00FF00 |
| 淺綠 | CCFFCC |
| 藍色 | 0000FF |
| 淺藍 | CCE5FF |
| 紫色 | 800080 |
| 灰色 | 808080 |
| 淺灰 | F2F2F2 |

---

## 🎯 最佳實踐

1. **定期備份**: 操作前備份重要檔案
2. **測試公式**: 設定公式後檢查計算結果
3. **一致格式**: 使用統一的格式標準
4. **批次操作**: 使用迴圈處理重複任務

---

更多範例請參考 [examples/](../examples/) 目錄。
