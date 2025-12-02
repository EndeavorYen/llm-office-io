# 系統設計文件 (System Design Document)

**專案名稱**: Office 文檔編輯工具集  
**版本**: 1.0.0  
**日期**: 2025-12-02  
**作者**: Development Team

---

## 目錄

1. [系統架構](#1-系統架構)
2. [模組設計](#2-模組設計)
3. [資料流程](#3-資料流程)
4. [介面設計](#4-介面設計)
5. [錯誤處理](#5-錯誤處理)
6. [擴展性設計](#6-擴展性設計)

---

## 1. 系統架構

### 1.1 整體架構

```
┌─────────────────────────────────────────────┐
│           使用者介面層                        │
│   (命令列 CLI / 自然語言解析器)                │
└─────────────────┬───────────────────────────┘
                  │
┌─────────────────▼───────────────────────────┐
│           應用程式層                          │
│                                             │
│  ┌──────────────┐    ┌──────────────┐      │
│  │ word_editor  │    │ ppt_editor   │      │
│  │  (Word編輯)  │    │  (PPT編輯)   │      │
│  └──────┬───────┘    └──────┬───────┘      │
│         │                   │              │
└─────────┼───────────────────┼──────────────┘
          │                   │
┌─────────▼───────────────────▼──────────────┐
│           函式庫層                           │
│                                             │
│  ┌──────────────┐    ┌──────────────┐      │
│  │ python-docx  │    │ python-pptx  │      │
│  └──────────────┘    └──────────────┘      │
└─────────────────┬───────────────────────────┘
                  │
┌─────────────────▼───────────────────────────┐
│           檔案系統層                          │
│       (.docx 檔案 / .pptx 檔案)              │
└─────────────────────────────────────────────┘
```

### 1.2 設計原則

1. **單一職責原則**: 每個模組只負責一項功能
2. **開放封閉原則**: 對擴展開放，對修改封閉
3. **依賴倒置原則**: 依賴抽象而非具體實現
4. **介面隔離原則**: 客戶端不應依賴它不需要的介面

---

## 2. 模組設計

### 2.1 Word 編輯器模組 (word_editor.py)

#### 2.1.1 類別架構

```python
class WordEditor:
    """Word 文檔編輯器核心類別"""
    
    def __init__(self, filepath: str)
    def save(self, output_path: Optional[str] = None)
    def list_structure(self)
    def replace_text(self, old_text: str, new_text: str, count: int = -1)
    def add_paragraph_after(self, search_text: str, new_content: str, heading_level: Optional[int] = None)
    def delete_paragraph(self, search_text: str)
    def insert_after_heading(self, heading_text: str, content: str, is_heading: bool = False, heading_level: int = 2)
    def add_bullet_points(self, heading_text: str, bullet_points: List[str])
```

#### 2.1.2 主要方法說明

| 方法 | 輸入 | 輸出 | 功能 |
|------|------|------|------|
| `list_structure()` | - | 文檔結構 | 列出所有段落索引 |
| `replace_text()` | 舊文字, 新文字, 次數 | 替換次數 | 替換文字 |
| `add_paragraph_after()` | 搜尋文字, 新內容 | bool | 在段落後添加 |
| `delete_paragraph()` | 搜尋文字 | bool | 刪除段落 |

#### 2.1.3 資料結構

```python
# Document 物件（來自 python-docx）
Document
├── paragraphs: List[Paragraph]
│   ├── text: str
│   ├── style: ParagraphStyle
│   └── runs: List[Run]
└── tables: List[Table]
    └── rows: List[Row]
        └── cells: List[Cell]
```

---

### 2.2 PowerPoint 編輯器模組 (ppt_editor.py)

#### 2.2.1 類別架構

```python
class PPTEditor:
    """PowerPoint 編輯器核心類別"""
    
    def __init__(self, filepath: str)
    def save(self, output_path: Optional[str] = None)
    def list_slides(self)
    def replace_text(self, old_text: str, new_text: str, slide_number: Optional[int] = None)
    def update_slide_title(self, slide_number: int, new_title: str)
    def add_text_to_slide(self, slide_number: int, text: str, position: str = 'body')
    def delete_slide(self, slide_number: int)
    def add_slide(self, title: str, layout_index: int = 1)
    def set_font(self, slide_number: int, font_name: str, font_size: Optional[int] = None)
    def get_slide_info(self, slide_number: int)
    def _get_slide_title(self, slide: Slide) -> Optional[str]
    def _get_slide_content_preview(self, slide: Slide) -> List[str]
```

#### 2.2.2 主要方法說明

| 方法 | 輸入 | 輸出 | 功能 |
|------|------|------|------|
| `list_slides()` | - | 投影片清單 | 列出所有投影片 |
| `replace_text()` | 舊文字, 新文字, 投影片編號 | 替換次數 | 替換文字 |
| `update_slide_title()` | 投影片編號, 新標題 | bool | 更新標題 |
| `add_slide()` | 標題, 版面編號 | Slide | 新增投影片 |
| `delete_slide()` | 投影片編號 | bool | 刪除投影片 |

#### 2.2.3 資料結構

```python
# Presentation 物件（來自 python-pptx）
Presentation
├── slides: List[Slide]
│   ├── shapes: List[Shape]
│   │   ├── title: Shape (TextFrame)
│   │   └── text_frame: TextFrame
│   │       └── paragraphs: List[Paragraph]
│   └── slide_layout: SlideLayout
└── slide_layouts: List[SlideLayout]
```

---

## 3. 資料流程

### 3.1 Word 文檔編輯流程

```
開始
  │
  ├─→ 讀取 .docx 檔案
  │     │
  │     └─→ 解析為 Document 物件
  │           │
  │           └─→ 執行編輯操作
  │                 │
  │                 ├─→ 遍歷 paragraphs
  │                 ├─→ 遍歷 tables
  │                 └─→ 修改內容
  │                       │
  ├─→ 儲存 Document 物件
  │     │
  │     └─→ 寫入 .docx 檔案
  │
結束
```

### 3.2 PowerPoint 編輯流程

```
開始
  │
  ├─→ 讀取 .pptx 檔案
  │     │
  │     └─→ 解析為 Presentation 物件
  │           │
  │           └─→ 執行編輯操作
  │                 │
  │                 ├─→ 遍歷 slides
  │                 ├─→ 遍歷 shapes
  │                 └─→ 修改內容
  │                       │
  ├─→ 儲存 Presentation 物件
  │     │
  │     └─→ 寫入 .pptx 檔案
  │
結束
```

### 3.3 命令執行流程

```
┌──────────────────┐
│ 接收命令列輸入    │
└────────┬─────────┘
         │
         ▼
┌──────────────────┐
│ 解析命令和參數    │
└────────┬─────────┘
         │
         ▼
┌──────────────────┐
│ 驗證輸入參數      │
└────────┬─────────┘
         │
         ├─→ 參數有誤 → 顯示錯誤訊息 → 結束
         │
         ▼
┌──────────────────┐
│ 載入文檔檔案      │
└────────┬─────────┘
         │
         ├─→ 檔案不存在 → 顯示錯誤訊息 → 結束
         │
         ▼
┌──────────────────┐
│ 執行編輯操作      │
└────────┬─────────┘
         │
         ├─→ 操作失敗 → 顯示錯誤訊息 → 結束
         │
         ▼
┌──────────────────┐
│ 儲存文檔          │
└────────┬─────────┘
         │
         ▼
┌──────────────────┐
│ 顯示成功訊息      │
└──────────────────┘
```

---

## 4. 介面設計

### 4.1 命令列介面 (CLI)

#### 4.1.1 參數解析器設計

使用 Python `argparse` 模組：

```python
parser = argparse.ArgumentParser(description='...')
parser.add_argument('file', help='檔案路徑')
parser.add_argument('--output', '-o', help='輸出路徑')

subparsers = parser.add_subparsers(dest='command')

# 添加子命令
replace_parser = subparsers.add_parser('replace')
replace_parser.add_argument('old', help='舊文字')
replace_parser.add_argument('new', help='新文字')
```

#### 4.1.2 輸出格式設計

**成功訊息格式**:
```
✓ [操作描述] ([統計資訊])
```

**錯誤訊息格式**:
```
✗ [錯誤描述]
建議: [解決方案]
```

**資訊訊息格式**:
```
[索引] 標題/內容預覽
```

---

## 5. 錯誤處理

### 5.1 錯誤類型

| 錯誤類型 | 錯誤碼 | 處理方式 |
|---------|--------|---------|
| 檔案不存在 | FILE_NOT_FOUND | 提示檔案路徑錯誤 |
| 格式錯誤 | INVALID_FORMAT | 提示檔案格式不支援 |
| 參數錯誤 | INVALID_ARGS | 顯示正確的參數格式 |
| 找不到目標 | NOT_FOUND | 提示目標不存在 |
| 權限不足 | PERMISSION_DENIED | 提示權限問題 |

### 5.2 錯誤處理策略

```python
try:
    # 執行操作
    result = perform_operation()
except FileNotFoundError:
    print(f"✗ 無法找到檔案: {filepath}")
    print("建議: 請檢查檔案路徑是否正確")
    sys.exit(1)
except PermissionError:
    print(f"✗ 權限不足，無法存取檔案")
    sys.exit(1)
except Exception as e:
    print(f"✗ 發生錯誤: {e}")
    import traceback
    traceback.print_exc()
    sys.exit(1)
```

### 5.3 輸入驗證

1. **檔案路徑驗證**: 檢查檔案是否存在
2. **格式驗證**: 檢查副檔名是否正確
3. **參數驗證**: 檢查必要參數是否提供
4. **範圍驗證**: 檢查數值是否在有效範圍內

---

## 6. 擴展性設計

### 6.1 外掛架構

未來可考慮實現外掛系統：

```python
class EditorPlugin:
    """編輯器外掛基礎類別"""
    
    def name(self) -> str:
        pass
    
    def execute(self, document, args):
        pass

class PluginManager:
    """外掛管理器"""
    
    def register_plugin(self, plugin: EditorPlugin):
        pass
    
    def execute_plugin(self, plugin_name: str, document, args):
        pass
```

### 6.2 支援其他檔案格式

可擴展支援：
- Excel (.xlsx) - 使用 openpyxl
- PDF - 使用 PyPDF2
- Markdown - 使用 markdown

### 6.3 批次處理

```python
class BatchProcessor:
    """批次處理器"""
    
    def __init__(self, file_list: List[str]):
        pass
    
    def apply_operation(self, operation: Callable):
        """對所有檔案套用操作"""
        pass
```

---

## 7. 效能考量

### 7.1 記憶體管理

- 使用串流處理大型文檔
- 及時釋放不需要的物件
- 避免一次載入整個文檔到記憶體

### 7.2 執行效能

- 只修改必要的部分
- 避免重複遍歷
- 使用快取儲存常用資訊

### 7.3 檔案 I/O 優化

- 使用緩衝區
- 批次寫入
- 只在必要時寫入磁碟

---

## 8. 安全性設計

### 8.1 輸入驗證

- 驗證所有使用者輸入
- 防止路徑遍歷攻擊
- 限制檔案大小

### 8.2 權限控制

- 只存取必要的檔案
- 不執行外部命令
- 不洩漏敏感資訊

### 8.3 資料保護

- 不記錄文檔內容
- 不上傳資料到外部
- 提供備份機制

---

## 9. 測試策略

### 9.1 單元測試

```python
class TestWordEditor(unittest.TestCase):
    def setUp(self):
        self.editor = WordEditor("test.docx")
    
    def test_replace_text(self):
        result = self.editor.replace_text("old", "new")
        self.assertGreater(result, 0)
    
    def tearDown(self):
        # 清理測試檔案
        pass
```

### 9.2 整合測試

- 測試完整的命令執行流程
- 驗證檔案輸出正確性
- 檢查錯誤處理機制

### 9.3 效能測試

- 測試不同大小的文檔
- 測量執行時間
- 監控記憶體使用

---

## 10. 部署架構

### 10.1 檔案結構

```
llm-office-io/
├── src/
│   ├── word_editor.py
│   ├── ppt_editor.py
│   └── __init__.py
├── docs/
│   ├── requirements.md
│   ├── design.md
│   └── user_manual.md
├── tests/
│   ├── test_word_editor.py
│   └── test_ppt_editor.py
├── examples/
│   ├── restructure_docx.py
│   └── enhance_docx.py
├── README.md
├── requirements.txt
└── setup.py
```

### 10.2 依賴管理

**requirements.txt**:
```
python-docx>=1.1.0
python-pptx>=1.0.0
```

### 10.3 安裝方式

```bash
pip install -r requirements.txt
```

---

## 附錄 A: 設計決策記錄

### A.1 為什麼使用命令列介面？

- 易於自動化
- 適合 AI 助理呼叫
- 跨平台相容性好

### A.2 為什麼不使用 GUI？

- 增加複雜度
- 不利於自動化
- 維護成本高

### A.3 為什麼使用 Python？

- 豐富的函式庫支援
- 易於開發和維護
- 跨平台支援良好

---

**版本歷史**

| 版本 | 日期 | 作者 | 變更描述 |
|------|------|------|----------|
| 1.0.0 | 2025-12-02 | Dev Team | 初版發布 |
