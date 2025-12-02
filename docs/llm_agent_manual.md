# LLM Agent 使用指南

## 📖 簡介

本工具集提供了專為 LLM Agent 設計的簡化 API，可以一步完成 Office 文檔的編輯操作。

---

## 🚀 快速開始

### 使用簡化 API

```python
from src.llm_api import replace_text, add_image, insert_table, batch_replace

# 1. 替換文字（自動判斷檔案類型）
result = replace_text(
    file_path="report.docx",
    old_text="2024",
    new_text="2025"
)

print(result)
# {
#     "success": True,
#     "operation": "replace_text",
#     "file_type": "word",
#     "result": {"count": 5},
#     "message": "成功替換 5 處",
#     "error": None
# }
```

---

## 📋 可用工具

### 1. replace_text - 文字替換

**用途**: 替換 Word/PowerPoint/Excel 文檔中的文字

```python
result = replace_text(
    file_path="document.docx",  # 支援 .docx, .pptx, .xlsx
    old_text="舊文字",
    new_text="新文字",
    output_path="output.docx"  # 可選
)
```

**返回格式**:
```json
{
  "success": true,
  "operation": "replace_text",
  "file_type": "word",
  "result": {"count": 3},
  "message": "成功替換 3 處",
  "error": null
}
```

---

### 2. add_image - 插入圖片

**用途**: 在 Word 或 PowerPoint 中插入圖片

```python
# Word 文檔
result = add_image(
    file_path="report.docx",
    image_path="logo.png",
    width_cm=5.0,
    position="第一章"  # 在包含「第一章」的段落後插入
)

# PowerPoint（需要指定投影片編號）
result = add_image(
    file_path="presentation.pptx",
    image_path="chart.png",
    slide_number=3,
    left_cm=5.0,
    top_cm=8.0,
    width_cm=15.0
)
```

---

### 3. insert_table - 插入表格

**用途**: 在 Word 文檔中插入表格

```python
# 空表格
result = insert_table(
    file_path="document.docx",
    rows=3,
    cols=4
)

# 帶數據的表格
data = [
    ["姓名", "年齡", "城市"],
    ["張三", "25", "台北"],
    ["李四", "30", "高雄"]
]

result = insert_table(
    file_path="document.docx",
    rows=3,
    cols=3,
    data=data,
    position="人員名單"  # 在包含此文字的段落後插入
)
```

---

### 4. batch_replace - 批次替換

**用途**: 一次處理多個檔案

```python
result = batch_replace(
    pattern="*.docx",           # 或 "reports/*.xlsx"
    old_text="2024",
    new_text="2025",
    recursive=True,             # 遞迴搜尋子目錄
    output_dir="updated/",      # 輸出到新目錄
    backup=True                 # 備份原檔案
)

print(result)
# {
#     "success": True,
#     "operation": "batch_replace",
#     "file_type": "mixed",
#     "result": {
#         "total": 15,
#         "success": 14,
#         "failed": 1,
#         "files": ["file1.docx", "file2.docx", ...]
#     },
#     "message": "處理 15 個檔案，成功 14 個",
#     "error": None
# }
```

---

## 🔧 通用接口

### execute_command

所有操作都可以通過統一接口調用：

```python
from src.llm_api import execute_command

result = execute_command(
    command="replace_text",
    file_path="doc.docx",
    old_text="A",
    new_text="B"
)
```

---

### JSON 模式（最適合 AI Agent）

```python
from src.llm_api import execute_json

# JSON 輸入
json_input = {
    "command": "replace_text",
    "params": {
        "file_path": "report.docx",
        "old_text": "2024",
        "new_text": "2025"
    }
}

# 獲取 JSON 輸出
result_json = execute_json(json_input)

# 或使用字符串
json_string = '{"command": "replace_text", "params": {"file_path": "test.docx", "old_text": "A", "new_text": "B"}}'
result_json = execute_json(json_string)
```

**返回 JSON 字符串**:
```json
{
  "success": true,
  "operation": "replace_text",
  "file_type": "word",
  "result": {"count": 2},
  "message": "成功替換 2 處",
  "error": null
}
```

---

## 📊 統一返回格式

所有函數都返回相同格式的字典：

```python
{
    "success": bool,        # 操作是否成功
    "operation": str,       # 操作名稱
    "file_type": str,       # 檔案類型 ("word", "ppt", "excel", "mixed")
    "result": dict,         # 操作結果（具體內容因操作而異）
    "message": str,         # 成功訊息
    "error": str | None     # 錯誤訊息（成功時為 None）
}
```

---

## 💡 使用範例

### 範例 1: 更新年度報告

```python
from src.llm_api import replace_text, add_image

# 1. 更新年份
result1 = replace_text("annual_report.docx", "2024", "2025")

if result1["success"]:
    # 2. 添加新的圖表
    result2 = add_image(
        "annual_report.docx",
        "2025_chart.png",
        width_cm=12.0,
        position="財務摘要"
    )
    
    if result2["success"]:
        print("報告更新完成！")
```

---

### 範例 2: 批次處理多個簡報

```python
from src.llm_api import batch_replace

result = batch_replace(
    pattern="presentations/*.pptx",
    old_text="Draft",
    new_text="Final",
    recursive=True,
    backup=True
)

print(f"處理結果: {result['result']['success']}/{result['result']['total']} 成功")
```

---

### 範例 3: 使用 JSON 接口

```python
from src.llm_api import execute_json
import json

# 定義多個操作
operations = [
    {
        "command": "replace_text",
        "params": {
            "file_path": "doc1.docx",
            "old_text": "A",
            "new_text": "B"
        }
    },
    {
        "command": "add_image",
        "params": {
            "file_path": "doc2.docx",
            "image_path": "logo.png",
            "width_cm": 5.0
        }
    }
]

# 執行所有操作
for op in operations:
    result_json = execute_json(op)
    result = json.loads(result_json)
    print(f"{result['operation']}: {result['message']}")
```

---

## ⚠️ 錯誤處理

所有函數都會捕捉異常並返回結構化錯誤：

```python
result = replace_text("nonexistent.docx", "A", "B")

# {
#     "success": False,
#     "operation": "replace_text",
#     "file_type": "unknown",
#     "result": None,
#     "message": "",
#     "error": "檔案不存在: nonexistent.docx"
# }

# 檢查並處理錯誤
if not result["success"]:
    print(f"錯誤: {result['error']}")
```

---

## 📖 工具描述檔案

完整的工具描述（JSON Schema 格式）可在以下檔案中找到：

- `docs/tool_descriptions.json` - 包含所有工具的參數定義和返回格式

這個檔案可以直接用於：
- LangChain tool definitions
- OpenAI function calling
- Anthropic Claude tools
- 其他 LLM framework

---

## 🎯 最佳實踐

1. **檢查返回值**: 始終檢查 `success` 欄位
2. **處理錯誤**: 當 `success=False` 時，檢查 `error` 欄位
3. **使用 JSON 模式**: 對於 LLM agents，推薦使用 `execute_json()`
4. **測試小範圍**: 先在單個檔案上測試，再批次處理

---

## 🔗 相關文檔

- [Word Editor 詳細指南](word_editor_guide.md)
- [PowerPoint Editor 指南](ppt_editor_guide.md)
- [Excel Editor 指南](excel_editor_guide.md)
- [批次處理指南](batch_processor_guide.md)

---

**更新日期**: 2025-12-02  
**版本**: 1.3.0
