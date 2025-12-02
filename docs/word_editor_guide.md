# Word Editor 完整使用指南

## 📖 簡介

Word Editor 提供 12 個強大功能，讓您輕鬆自動化 Word 文檔的編輯操作。

---

## 🚀 快速開始

```python
from src.word_editor import WordEditor

# 開啟文檔
editor = WordEditor("document.docx")

# 執行操作
editor.replace_text("舊文字", "新文字")
editor.save("output.docx")
```

---

## 📋 功能列表

### 1. 文字替換 `replace_text()`

```python
# 替換所有出現的文字
count = editor.replace_text("2024", "2025")

# 只替換前 3 次
count = editor.replace_text("Apple", "Orange", count=3)
```

### 2. 圖片插入 `add_image()` 🆕

```python
# 在文檔末尾插入圖片
editor.add_image("photo.jpg", width_cm=12.0)

# 在特定位置後插入
editor.add_image("logo.png", width_cm=5.0, position="第一章")
```

**參數**:
- `image_path`: 圖片檔案路徑
- `width_cm`: 圖片寬度（公分），預設 10.0
- `position`: 插入位置，None 表示文檔末尾

---

### 3. 表格插入 `insert_table()` 🆕

```python
# 插入 3x4 空表格
editor.insert_table(rows=3, cols=4)

# 插入表格並填充數據
data = [
    ["姓名", "年齡", "城市"],
    ["張三", "25", "台北"],
    ["李四", "30", "高雄"]
]
editor.insert_table(rows=3, cols=3, data=data)

# 在特定位置後插入
editor.insert_table(rows=2, cols=3, position="總結")
```

---

### 4. 更新表格儲存格 `update_table_cell()` 🆕

```python
# 更新第 1 個表格的第 0 行第 1 列
editor.update_table_cell(
    table_index=0,  # 第 1 個表格
    row=0,          # 第 1 行
    col=1,          # 第 2 列
    text="已更新"
)
```

---

### 5. 段落格式設定 `set_paragraph_format()` 🆕

```python
# 設定包含「標題」的段落為粗體、18pt、置中
editor.set_paragraph_format(
    search_text="標題",
    font_size=18,
    bold=True,
    alignment="center"
)

# 設定斜體
editor.set_paragraph_format(
    search_text="重要說明",
    italic=True,
    alignment="justify"
)
```

**對齊選項**: `'left'`, `'center'`, `'right'`, `'justify'`

---

### 6. 插入分頁符號 `add_page_break()` 🆕

```python
# 在文檔末尾插入分頁
editor.add_page_break()

# 在特定文字後插入分頁
editor.add_page_break(after_text="第一章結束")
```

---

### 7. 段落刪除 `delete_paragraph()`

```python
# 刪除包含特定文字的段落
editor.delete_paragraph("待刪除的內容")
```

---

### 8. 新增段落 `add_paragraph_after()`

```python
# 在特定段落後添加普通段落
editor.add_paragraph_after(
    search_text="序言",
    new_content="這是新增的內容"
)

# 添加標題段落
editor.add_paragraph_after(
    search_text="第一章",
    new_content="新的小節",
    heading_level=2
)
```

---

### 9. 列出文檔結構 `list_structure()`

```python
# 顯示所有標題和段落
editor.list_structure()
```

輸出範例:
```
=== 文檔結構 ===

[0] 📌 第一章：簡介
[1]    這是第一章的內容...
[2] 📌 第二章：方法
[3]    研究方法包括...
```

---

### 10. 在標題後插入內容 `insert_after_heading()`

```python
# 在「第一章」後插入段落
editor.insert_after_heading(
    heading_text="第一章",
    content="這是新增的段落"
)

# 插入子標題
editor.insert_after_heading(
    heading_text="第一章",
    content="1.1 背景",
    is_heading=True,
    heading_level=2
)
```

---

### 11. 添加項目符號 `add_bullet_points()`

```python
# 在標題後添加多個項目
bullets = [
    "第一個要點",
    "第二個要點",
    "第三個要點"
]
editor.add_bullet_points("總結", bullets)
```

---

### 12. 儲存文檔 `save()`

```python
# 覆蓋原檔案
editor.save()

# 另存新檔
editor.save("new_document.docx")
```

---

## 💡 實用範例

### 範例 1: 年度報告更新

```python
editor = WordEditor("annual_report_2024.docx")

# 更新年份
editor.replace_text("2024", "2025")

# 插入新章節
editor.add_page_break(after_text="第三章結束")
editor.add_paragraph_after(
    search_text="第三章結束",
    new_content="第四章：未來展望",
    heading_level=1
)

# 添加內容
bullets = ["擴大市場", "提升品質", "數位轉型"]
editor.add_bullet_points("第四章：未來展望", bullets)

editor.save("annual_report_2025.docx")
```

---

### 範例 2: 添加公司標誌

```python
editor = WordEditor("proposal.docx")

# 在標題後插入標誌
editor.add_image(
    "company_logo.png",
    width_cm=5.0,
    position="提案書"
)

# 設定標題格式
editor.set_paragraph_format(
    search_text="提案書",
    font_size=24,
    bold=True,
    alignment="center"
)

editor.save()
```

---

### 範例 3: 創建報告表格

```python
editor = WordEditor("report.docx")

# 插入數據表格
data = [
    ["項目", "Q1", "Q2", "Q3", "Q4"],
    ["營收", "100M", "120M", "115M", "140M"],
    ["成本", "60M", "70M", "65M", "75M"]
]
editor.insert_table(rows=3, cols=5, data=data, position="財務摘要")

# 更新特定儲存格
editor.update_table_cell(0, 0, 0, "財務項目")

editor.save()
```

---

## ⚠️ 注意事項

1. **檔案格式**: 僅支援 `.docx` 格式
2. **備份建議**: 操作前建議備份原檔案
3. **索引從 0 開始**: 表格索引、行列索引都從 0 開始
4. **圖片格式**: 支援 JPG、PNG 等常見格式

---

## 🎯 最佳實踐

1. **先測試**: 在少量文檔上測試腳本
2. **使用版本控制**: 為重要文檔啟用版本控制
3. **檢查結果**: 操作後檢查輸出文檔
4. **批次處理**: 使用 batch_processor 處理多個檔案

---

更多範例請參考 [examples/](../examples/) 目錄。
