# Office 文檔編輯工具集

> 強大的 Word、PowerPoint 和 Excel 命令列編輯工具  
> 支援自然語言指令和批次處理

[![Python Version](https://img.shields.io/badge/python-3.8+-blue.svg)](https://www.python.org/downloads/)
[![License](https://img.shields.io/badge/license-MIT-green.svg)](LICENSE)

---

## 📖 簡介

這是一套專為自動化文檔編輯而設計的命令列工具，支援：

- ✅ Word 文檔 (.docx) 編輯
- ✅ PowerPoint (.pptx) 編輯  
- ✅ Excel (.xlsx) 編輯 ✨
- ✅ 批次文字替換
- ✅ 內容管理和結構調整
- ✅ AI 助理友好的介面設計

**適用對象**: 開發人員、技術文檔編輯者、內容管理人員、AI 助理

---

## 🚀 快速開始

### 安裝

```bash
# 1. 克隆或下載專案
cd llm-office-io

# 2. 安裝依賴
pip install -r requirements.txt

# 3. 驗證安裝
python src/word_editor.py --help
python src/ppt_editor.py --help
```

### 快速範例

```bash
# Word 文檔：替換所有「2024」為「2025」
python src/word_editor.py report.docx replace "2024" "2025"

# PowerPoint：更新第一張投影片的標題
python src/ppt_editor.py slides.pptx update-title 1 "新標題"

# Excel：替換所有工作表中的文字
python src/excel_editor.py data.xlsx replace "舊值" "新值"

# 查看文檔結構
python src/word_editor.py document.docx list
python src/ppt_editor.py presentation.pptx list
python src/excel_editor.py workbook.xlsx list
```

---

## 📂 專案結構

```
llm-office-io/
├── src/                    # 源代碼
│   ├── word_editor.py      # Word 編輯器
│   ├── ppt_editor.py       # PowerPoint 編輯器
│   ├── excel_editor.py     # Excel 編輯器 ✨
│   ├── constants.py        # 常量定義
│   ├── __init__.py         # 套件初始化
│   └── read_docx.py        # Word 讀取工具
│
├── docs/                   # 文檔
│   ├── requirements.md     # 需求規格書
│   ├── design.md          # 系統設計文件
│   ├── user_manual.md     # 使用說明
│   ├── excel_editor_guide.md # Excel 編輯器指南 ✨
│   └── llm_agent_manual.md # AI 助理手冊
│
├── examples/              # 範例腳本
│   ├── restructure_docx.py # 文檔重構範例
│   └── enhance_docx.py     # 文檔增強範例
│
├── tests/                 # 測試檔案
│   ├── test_word_editor.py
│   ├── test_ppt_editor.py
│   └── test_excel_editor.py ✨
│
├── README.md              # 本文件
├── requirements.txt       # Python 依賴
└── .gitignore
```

---

## 🛠️ 功能特色

### Word 編輯器 (word_editor.py)

| 功能 | 命令 | 說明 |
|------|------|------|
| 查看結構 | `list` | 列出所有段落和標題 |
| 替換文字 | `replace` | 批量替換文字內容 |
| 插入內容 | `insert-after-heading` | 在標題後插入新內容 |
| 添加列表 | `add-bullets` | 添加項目符號列表 |
| 刪除段落 | `delete` | 刪除指定段落 |

### PowerPoint 編輯器 (ppt_editor.py)

| 功能 | 命令 | 說明 |
|------|------|------|
| 列出投影片 | `list` | 顯示所有投影片 |
| 替換文字 | `replace` | 批量替換文字 |
| 更新標題 | `update-title` | 修改投影片標題 |
| 新增投影片 | `add-slide` | 添加新投影片 |
| 刪除投影片 | `delete-slide` | 移除投影片 |
| 設定字體 | `set-font` | 修改字體樣式 |

### Excel 編輯器 (excel_editor.py) ✨

| 功能 | 命令 | 說明 |
|------|------|------|
| 列出工作表 | `list` | 顯示所有工作表 |
| 查看內容 | `view` | 查看工作表資料 |
| 替換文字 | `replace` | 批量替換文字 |
| 更新儲存格 | `update-cell` | 修改儲存格值 |
| 新增行 | `add-row` | 插入新資料行 |
| 刪除行 | `delete-row` | 移除資料行 |
| 搜尋儲存格 | `find` | 搜尋特定文字 |

---

## 📚 使用文檔

- **[使用說明](docs/user_manual.md)** - 完整的使用手冊，包含範例和 FAQ
- **[需求規格](docs/requirements.md)** - 系統需求和功能規格
- **[設計文件](docs/design.md)** - 技術架構和設計決策
- **[AI 助理手冊](docs/llm_agent_manual.md)** - 給 LLM Agent 的詳細操作指南
- **[PPT 編輯器指南](docs/ppt_editor_guide.md)** - PowerPoint 編輯器快速參考

---

## 💡 常用範例

### 範例 1：更新年度報告

```bash
# 1. 查看文檔結構
python src/word_editor.py annual_report.docx list

# 2. 批量更新年份
python src/word_editor.py annual_report.docx replace "2024" "2025"

# 3. 更新講師名稱
python src/word_editor.py annual_report.docx replace "John" "Sarah"

# 4. 另存新檔
python src/word_editor.py annual_report.docx replace "Draft" "Final" --output final_report.docx
```

### 範例 2：簡報批次處理

```bash
# 1. 列出所有投影片
python src/ppt_editor.py training.pptx list

# 2. 替換整份簡報的文字
python src/ppt_editor.py training.pptx replace "舊版本" "新版本"

# 3. 更新封面標題
python src/ppt_editor.py training.pptx update-title 1 "2025 培訓課程"

# 4. 新增結束頁
python src/ppt_editor.py training.pptx add-slide "Q&A"
```

### 範例 3：文檔結構調整

```bash
# 在特定標題後添加內容
python src/word_editor.py doc.docx insert-after-heading "摘要" "本文檔更新於 2025 年"

# 添加功能列表
python src/word_editor.py doc.docx add-bullets "主要功能" "功能A" "功能B" "功能C"

# 刪除過時內容
python src/word_editor.py doc.docx delete "待刪除"
```

---

## ⚙️ 系統需求

- **Python**: 3.8 或更新版本
- **作業系統**: Windows / Linux / macOS
- **依賴套件**: python-docx, python-pptx

---

## 📋 功能路線圖

### ✅ 已完成 (v1.1)
- [x] Word 文檔基本編輯
- [x] PowerPoint 基本編輯
- [x] Excel 基本編輯 ✨
- [x] 命令列介面
- [x] 完整文檔
- [x] 單元測試架構

### 🚧 計劃中 (v2.0)
- [ ] 批次處理模式
- [ ] 圖片和圖表操作
- [ ] 配置檔支援
- [ ] GUI 介面（可選）

---

## 🤝 貢獻

歡迎提交 Issue 和 Pull Request！

### 開發設置

```bash
# 1. Fork 專案
# 2. 創建功能分支
git checkout -b feature/your-feature

# 3. 提交變更
git commit -m "Add some feature"

# 4. 推送到分支
git push origin feature/your-feature

# 5. 創建 Pull Request
```

---

## ⚠️ 注意事項

1. **備份重要文件** - 修改前建議先備份
2. **測試環境** - 先在測試文件上驗證命令
3. **編碼問題** - 確保終端支援 UTF-8
4. **檔案格式** - 僅支援 Office 2007+ (.docx/.pptx)

---

## 📞 支援與回饋

- 📖 查看 [使用說明](docs/user_manual.md)
- 📧 聯絡開發團隊
- 🐛 [回報問題](../../issues)

---

## 📄 授權

本專案採用 MIT 授權 - 詳見 [LICENSE](LICENSE) 檔案

---

## 🙏 致謝

- [python-docx](https://python-docx.readthedocs.io/) - Word 文檔處理
- [python-pptx](https://python-pptx.readthedocs.io/) - PowerPoint 處理
- [openpyxl](https://openpyxl.readthedocs.io/) - Excel 處理 ✨

---

**最後更新**: 2025-12-02  
**版本**: 1.1.0
