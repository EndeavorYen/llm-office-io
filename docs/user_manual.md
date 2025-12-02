# 使用說明文件 (User Manual)

**專案名稱**: Office 文檔編輯工具集  
**版本**: 1.0.0  
**日期**: 2025-12-02

---

## 目錄

1. [快速開始](#1-快速開始)
2. [Word 編輯器使用手冊](#2-word-編輯器使用手冊)
3. [PowerPoint 編輯器使用手冊](#3-powerpoint-編輯器使用手冊)
4. [常見問題 (FAQ)](#4-常見問題-faq)
5. [最佳實踐](#5-最佳實踐)
6. [疑難排解](#6-疑難排解)

---

## 1. 快速開始

### 1.1 安裝

#### 系統需求
- Python 3.8 或更新版本
- Windows / Linux / macOS

#### 安裝依賴

```bash
cd llm-office-io
pip install -r requirements.txt
```

#### 驗證安裝

```bash
python src/word_editor.py --help
python src/ppt_editor.py --help
```

---

## 2. Word 編輯器使用手冊

### 2.1 基本語法

```bash
python src/word_editor.py <文件.docx> <命令> [參數] [選項]
```

### 2.2 命令清單

#### 2.2.1 list - 查看文檔結構

**功能**: 列出文檔中所有段落和標題

**語法**:
```bash
python src/word_editor.py document.docx list
```

**範例輸出**:
```
=== 文檔結構 ===

[0] 📌 第一章 標題
[1]    這是第一段內容...
[2] 📌 第二章 標題
[3]    這是第二段內容...
```

**使用場景**:
- 不知道文檔內容時
- 需要找到特定段落的位置
- 確認文檔結構

---

#### 2.2.2 replace - 替換文字

**功能**: 在文檔中批量替換文字

**語法**:
```bash
python src/word_editor.py document.docx replace "舊文字" "新文字" [--count N]
```

**參數說明**:
- `old`: 要替換的文字
- `new`: 新文字
- `--count N`: (可選) 限制替換次數，-1 表示全部

**範例**:

```bash
# 範例 1: 替換所有出現的文字
python src/word_editor.py report.docx replace "2024" "2025"
# 輸出: ✓ 已替換 15 處「2024」→「2025」

# 範例 2: 只替換前 3 處
python src/word_editor.py report.docx replace "錯字" "正確字" --count 3
# 輸出: ✓ 已替換 3 處「錯字」→「正確字」

# 範例 3: 替換中文標點
python src/word_editor.py doc.docx replace "，" "、"
```

**注意事項**:
- 區分大小寫
- 會搜尋整份文檔（包括表格）
- 保留原有格式

---

#### 2.2.3 insert-after-heading - 在標題後插入內容

**功能**: 在特定標題後插入新內容

**語法**:
```bash
python src/word_editor.py document.docx insert-after-heading "標題文字" "內容" [--is-heading] [--heading-level N]
```

**參數說明**:
- `heading`: 標題文字
- `content`: 要插入的內容
- `--is-heading`: (可選) 插入的內容也是標題
- `--heading-level N`: (可選) 標題層級（1-3），預設為 2

**範例**:

```bash
# 範例 1: 在標題後加普通段落
python src/word_editor.py doc.docx insert-after-heading "系統概述" "這是補充說明"
# 輸出: ✓ 已在標題「系統概述」後插入內容

# 範例 2: 在標題後加子標題
python src/word_editor.py doc.docx insert-after-heading "第一章" "1.1 簡介" --is-heading --heading-level 2
# 輸出: ✓ 已在標題「第一章」後插入內容

# 範例 3: 添加多行內容
python src/word_editor.py doc.docx insert-after-heading "摘要" "第一行\n第二行\n第三行"
```

---

#### 2.2.4 add-bullets - 添加項目符號

**功能**: 在標題後批量添加項目符號列表

**語法**:
```bash
python src/word_editor.py document.docx add-bullets "標題" "項目1" "項目2" "項目3" ...
```

**範例**:

```bash
# 範例 1: 添加功能列表
python src/word_editor.py doc.docx add-bullets "主要功能" "功能A" "功能B" "功能C"
# 輸出: ✓ 已在「主要功能」後添加 3 個項目

# 範例 2: 添加需求清單
python src/word_editor.py spec.docx add-bullets "必備技能" "Python 基礎" "Git 使用" "Linux 操作"
```

---

#### 2.2.5 delete - 刪除段落

**功能**: 刪除包含特定文字的段落

**語法**:
```bash
python src/word_editor.py document.docx delete "搜尋文字"
```

**範例**:

```bash
# 範例 1: 刪除過時內容
python src/word_editor.py doc.docx delete "待刪除"
# 輸出: ✓ 已刪除段落: 待刪除...

# 範例 2: 刪除特定章節
python src/word_editor.py doc.docx delete "舊版說明"
```

**⚠️ 警告**: 刪除操作不可逆，建議先備份

---

#### 2.2.6 通用選項

**--output / -o**: 另存新檔

```bash
python src/word_editor.py original.docx replace "test" "prod" --output final.docx
```

### 2.3 完整範例

#### 範例 1: 更新培訓文檔

```bash
# 步驟 1: 查看文檔結構
python src/word_editor.py training.docx list

# 步驟 2: 更新年份
python src/word_editor.py training.docx replace "2024" "2025"

# 步驟 3: 更新講師名稱
python src/word_editor.py training.docx replace "John" "Sarah"

# 步驟 4: 在摘要後添加說明
python src/word_editor.py training.docx insert-after-heading "摘要" "本課程已更新至 2025 版本"

# 步驟 5: 添加新功能列表
python src/word_editor.py training.docx add-bullets "新功能" "功能1" "功能2" "功能3"
```

---

## 3. PowerPoint 編輯器使用手冊

### 3.1 基基本語法

```bash
python src/ppt_editor.py <簡報.pptx> <命令> [參數] [選項]
```

### 3.2 命令清單

#### 3.2.1 list - 列出投影片

**功能**: 顯示所有投影片的標題和內容摘要

**語法**:
```bash
python src/ppt_editor.py presentation.pptx list
```

**範例輸出**:
```
=== 簡報結構 (共 10 張投影片) ===

[投影片 1] 封面標題
  • 副標題文字...
  • 作者資訊...

[投影片 2] 第一章
  • 內容摘要第一行...
```

---

#### 3.2.2 replace - 替換文字

**功能**: 替換投影片中的文字

**語法**:
```bash
python src/ppt_editor.py presentation.pptx replace "舊文字" "新文字" [--slide N]
```

**範例**:

```bash
# 範例 1: 替換整份簡報
python src/ppt_editor.py slides.pptx replace "2024" "2025"
# 輸出: ✓ 在所有投影片中替換了 25 處「2024」→「2025」

# 範例 2: 只替換第 3 張
python src/ppt_editor.py slides.pptx replace "講師A" "講師B" --slide 3
# 輸出: ✓ 在投影片 3中替換了 2 處「講師A」→「講師B」
```

---

#### 3.2.3 update-title - 更新投影片標題

**功能**: 修改指定投影片的標題

**語法**:
```bash
python src/ppt_editor.py presentation.pptx update-title <投影片編號> "新標題"
```

**範例**:

```bash
# 範例 1: 更新第 1 張標題
python src/ppt_editor.py slides.pptx update-title 1 "2025 年度計畫"
# 輸出: ✓ 投影片 1 標題已更新
#        舊標題: 2024 年度計畫
#        新標題: 2025 年度計畫

# 範例 2: 更新章節標題
python src/ppt_editor.py slides.pptx update-title 5 "系統架構介紹"
```

---

#### 3.2.4 add-slide - 新增投影片

**功能**: 在簡報最後新增投影片

**語法**:
```bash
python src/ppt_editor.py presentation.pptx add-slide "標題" [--layout N]
```

**範例**:

```bash
# 範例 1: 新增結束頁
python src/ppt_editor.py slides.pptx add-slide "Q&A"
# 輸出: ✓ 已新增投影片 11: Q&A

# 範例 2: 新增感謝頁
python src/ppt_editor.py slides.pptx add-slide "感謝聆聽"
```

---

#### 3.2.5 delete-slide - 刪除投影片

**功能**: 刪除指定投影片

**語法**:
```bash
python src/ppt_editor.py presentation.pptx delete-slide <投影片編號>
```

**範例**:

```bash
python src/ppt_editor.py slides.pptx delete-slide 5
# 輸出: ✓ 已刪除投影片 5: 舊章節標題
```

**⚠️ 注意**: 刪除後，後續投影片編號會自動前移

---

#### 3.2.6 set-font - 設定字體

**功能**: 修改投影片的字體和大小

**語法**:
```bash
python src/ppt_editor.py presentation.pptx set-font <投影片編號> "字體名稱" [--size N]
```

**範例**:

```bash
# 範例 1: 只改字體
python src/ppt_editor.py slides.pptx set-font 1 "微軟正黑體"
# 輸出: ✓ 投影片 1 已更新字體: 微軟正黑體

# 範例 2: 改字體和大小
python src/ppt_editor.py slides.pptx set-font 1 "微軟正黑體" --size 24
# 輸出: ✓ 投影片 1 已更新字體: 微軟正黑體 (24pt)
```

---

#### 3.2.7 info - 查看投影片詳情

**功能**: 顯示投影片的詳細內容

**語法**:
```bash
python src/ppt_editor.py presentation.pptx info <投影片編號>
```

**範例**:

```bash
python src/ppt_editor.py slides.pptx info 3
# 輸出: 顯示投影片 3 的所有文字內容和形狀資訊
```

---

### 3.3 完整範例

#### 範例 1: 更新年度簡報

```bash
# 步驟 1: 查看簡報
python src/ppt_editor.py annual_report.pptx list

# 步驟 2: 替換年份
python src/ppt_editor.py annual_report.pptx replace "2024" "2025"

# 步驟 3: 更新封面標題
python src/ppt_editor.py annual_report.pptx update-title 1 "2025 年度報告"

# 步驟 4: 新增結束頁
python src/ppt_editor.py annual_report.pptx add-slide "謝謝"
```

---

## 4. 常見問題 (FAQ)

### Q1: 工具支援哪些檔案格式？

**A**: 
- Word: `.docx` (Office 2007+)
- PowerPoint: `.pptx` (Office 2007+)
- 不支援舊版 `.doc` 和 `.ppt` 格式

### Q2: 會不會覆蓋原檔案？

**A**: 預設會覆蓋原檔案。建議：
1. 操作前先備份
2. 使用 `--output` 另存新檔

### Q3: 替換文字時區分大小寫嗎？

**A**: 是的，完全區分大小寫。"Test" 和 "test" 會被視為不同的文字。

### Q4: 如何批次處理多個檔案？

**A**: 可以使用 shell 腳本：

```bash
# Windows (PowerShell)
Get-ChildItem *.docx | ForEach-Object { python src/word_editor.py $_.Name replace "old" "new" }

# Linux/macOS
for file in *.docx; do python src/word_editor.py "$file" replace "old" "new"; done
```

### Q5: 修改後格式會跑掉嗎？

**A**: 大部分格式會保留，包括：
- 字體、顏色、大小
- 粗體、斜體、底線
- 段落對齊
- 表格結構

但複雜的排版效果可能有些許差異。

### Q6: 可以撤銷操作嗎？

**A**: 工具本身不支援撤銷。建議：
1. 使用 `--output` 另存新檔
2. 使用版本控制系統 (Git)
3. 定期備份

### Q7: 命令執行很慢怎麼辦？

**A**: 
- 小型文檔 (< 50 頁) 應該在 2 秒內完成
- 大型文檔可能需要更長時間
- 如果超過 30 秒，可能是文檔損壞

### Q8: 找不到目標文字怎麼辦？

**A**: 
1. 先執行 `list` 命令查看內容
2. 檢查文字拼寫和大小寫
3. 確認目標文字確實存在

---

## 5. 最佳實踐

### 5.1 安全操作

✅ **建議做法**:
- 使用 `--output` 另存新檔
- 定期備份重要文件
- 先在測試文檔上實驗

❌ **避免做法**:
- 在沒有備份的情況下直接修改原檔
- 同時對多個重要文件進行批次操作
- 忽略錯誤訊息繼續操作

### 5.2 工作流程

**推薦流程**:

```
1. 查看結構 (list)
   ↓
2. 確認操作目標
   ↓
3. 先用 --output 測試
   ↓
4. 檢查結果
   ↓
5. 確認無誤後執行正式操作
```

### 5.3 命名規範

**良好的檔案命名**:
```
✓ report_2024_v1.docx
✓ presentation_final.pptx
✓ user_manual_zh.docx

✗ 文件.docx
✗ 新增 文字文件.docx
✗ report final (1).docx
```

### 5.4 版本管理

```bash
# 使用日期標記版本
python src/word_editor.py report.docx replace "old" "new" --output report_20251202.docx

# 使用版本號
python src/word_editor.py doc.docx replace "v1.0" "v1.1" --output doc_v1.1.docx
```

---

## 6. 疑難排解

### 6.1 常見錯誤及解決方案

#### 錯誤 1: `FileNotFoundError`

```
✗ 無法找到檔案: document.docx
```

**解決方案**:
1. 檢查檔案路徑是否正確
2. 確認檔案確實存在
3. 使用絕對路徑

```bash
# 使用絕對路徑
python src/word_editor.py "C:\Users\...\document.docx" list
```

#### 錯誤 2: `PermissionError`

```
✗ 權限不足，無法存取檔案
```

**解決方案**:
1. 關閉正在開啟該檔案的程式 (如 Word, PowerPoint)
2. 檢查檔案權限
3. 以管理員身份執行

#### 錯誤 3: 編碼錯誤

```
UnicodeDecodeError: ...
```

**解決方案**:
1. 確保終端支援 UTF-8
2. Windows 用戶設定編碼：
```powershell
$OutputEncoding = [System.Text.Encoding]::UTF8
```

#### 錯誤 4: 找不到目標

```
✗ 找不到「XXX」
```

**解決方案**:
1. 執行 `list` 確認內容
2. 檢查大小寫
3. 檢查是否有多餘空格

### 6.2 效能問題

**問題**: 執行很慢

**解決方案**:
1. 檢查文檔大小
2. 關閉不必要的程式
3. 分段處理大型文檔

### 6.3 取得幫助

```bash
# 查看命令說明
python src/word_editor.py --help
python src/ppt_editor.py --help

# 查看特定命令說明
python src/word_editor.py document.docx replace --help
```

---

## 附錄 A: 命令速查表

### Word 編輯器

| 命令 | 簡短說明 |
|------|---------|
| `list` | 查看結構 |
| `replace "A" "B"` | 替換文字 |
| `insert-after-heading "標題" "內容"` | 標題後插入 |
| `add-bullets "標題" "1" "2"` | 添加列表 |
| `delete "文字"` | 刪除段落 |

### PowerPoint 編輯器

| 命令 | 簡短說明 |
|------|---------|
| `list` | 列出投影片 |
| `replace "A" "B"` | 替換文字 |
| `update-title N "標題"` | 更新標題 |
| `add-slide "標題"` | 新增投影片 |
| `delete-slide N` | 刪除投影片 |
| `set-font N "字體"` | 設定字體 |
| `info N` | 查看詳情 |

---

## 附錄 B: 快捷鍵盤

```bash
# 建立常用命令別名 (Linux/macOS)
alias wed="python /path/to/src/word_editor.py"
alias pped="python /path/to/src/ppt_editor.py"

# 使用別名
wed document.docx list
pped slides.pptx replace "old" "new"
```

---

**需要更多幫助？**

- 查看 [設計文件](design.md) 了解技術細節
- 查看 [需求規格](requirements.md) 了解功能範圍
- 查看 [LLM Agent 手冊](../LLM_AGENT_MANUAL.md) 了解 AI 使用方式

---

**版本歷史**

| 版本 | 日期 | 變更 |
|------|------|------|
| 1.0.0 | 2025-12-02 | 初版發布 |
