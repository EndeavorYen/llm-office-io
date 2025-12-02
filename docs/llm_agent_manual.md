# Office 文檔編輯工具 - LLM Agent 使用手冊

> **目標讀者**: LLM Agent (AI Assistant)  
> **用途**: 在與用戶互動時，根據自然語言需求操作 Word 和 PowerPoint 文檔  
> **更新日期**: 2025-12-01

---

## 📚 目錄

1. [工具概覽](#工具概覽)
2. [Word 編輯器詳解](#word-編輯器詳解)
3. [PowerPoint 編輯器詳解](#powerpoint-編輯器詳解)
4. [自然語言到命令的轉換](#自然語言到命令的轉換)
5. [最佳實踐](#最佳實踐)
6. [錯誤處理](#錯誤處理)
7. [完整工作流程範例](#完整工作流程範例)

---

## 工具概覽

### 可用工具

| 工具 | 檔案 | 支援格式 | 主要功能 |
|------|------|---------|---------|
| Word 編輯器 | `word_editor.py` | `.docx` | 文字替換、段落管理、內容插入 |
| PowerPoint 編輯器 | `ppt_editor.py` | `.pptx` | 投影片管理、文字替換、標題更新 |

### 核心原則

1. **先查看再修改**: 使用 `list` 命令了解文檔結構
2. **明確指定**: 盡可能指定具體位置（投影片編號、標題等）
3. **備份建議**: 建議用戶備份或使用 `--output` 另存新檔
4. **驗證結果**: 執行命令後確認輸出訊息

---

## Word 編輯器詳解

### 基本語法

```bash
python word_editor.py <檔案路徑> <命令> [參數] [選項]
```

### 命令清單

#### 1. `list` - 查看文檔結構

**用途**: 列出文檔中所有段落和標題的索引

**語法**:
```bash
python word_editor.py document.docx list
```

**何時使用**:
- 用戶要求修改內容但未指定位置
- 需要了解文檔結構
- 尋找特定內容的位置

**範例對話**:
```
用戶: "把第三個標題改成..."
Agent: 先執行 list 查看哪個是第三個標題
```

---

#### 2. `replace` - 替換文字

**用途**: 在文檔中批量替換文字（支援段落和表格）

**語法**:
```bash
python word_editor.py document.docx replace "舊文字" "新文字" [--count N]
```

**參數**:
- `old`: 要替換的文字（完全匹配）
- `new`: 新文字
- `--count N`: (可選) 限制替換次數，-1 表示全部

**使用時機**:
- 用戶說「把所有...改成...」
- 批量更新術語、名稱、日期
- 更正錯字或統一用詞

**範例**:
```bash
# 替換所有出現的文字
python word_editor.py report.docx replace "2024" "2025"

# 只替換第一處
python word_editor.py report.docx replace "錯字" "正確字" --count 1
```

**注意事項**:
- 區分大小寫
- 會搜尋整份文檔（包括表格）
- 保留原有格式

---

#### 3. `add-after` - 在段落後添加內容

**用途**: 在包含特定文字的段落後插入新內容

**語法**:
```bash
python word_editor.py document.docx add-after "搜尋文字" "新內容" [--heading N]
```

**參數**:
- `search`: 搜尋關鍵字（找到第一個匹配的段落）
- `content`: 要添加的內容
- `--heading N`: (可選) 作為 N 級標題插入

**使用時機**:
- 「在...後面加上...」
- 在特定章節後補充內容
- 插入新段落

**範例**:
```bash
# 在包含「摘要」的段落後加普通文字
python word_editor.py doc.docx add-after "摘要" "這是補充說明"

# 插入標題
python word_editor.py doc.docx add-after "第一章" "新章節" --heading 2
```

---

#### 4. `insert-after-heading` - 在標題後插入

**用途**: 在特定標題後插入內容（更精確）

**語法**:
```bash
python word_editor.py document.docx insert-after-heading "標題文字" "內容" [--is-heading] [--heading-level N]
```

**參數**:
- `heading`: 標題文字
- `content`: 要插入的內容
- `--is-heading`: 插入的內容也是標題
- `--heading-level N`: 標題層級（1-3）

**使用時機**:
- 用戶明確指定在某個標題後插入
- 需要在章節結構中添加內容

**範例**:
```bash
# 在「系統概述」標題後加普通段落
python word_editor.py doc.docx insert-after-heading "系統概述" "這是詳細說明"

# 在標題後加子標題
python word_editor.py doc.docx insert-after-heading "第一章" "1.1 簡介" --is-heading --heading-level 2
```

---

#### 5. `delete` - 刪除段落

**用途**: 刪除包含特定文字的段落

**語法**:
```bash
python word_editor.py document.docx delete "搜尋文字"
```

**使用時機**:
- 「刪除...這段」
- 移除過時內容
- 清理文檔

**範例**:
```bash
python word_editor.py doc.docx delete "待刪除的內容"
```

**警告**: 不可逆操作，建議先備份

---

#### 6. `add-bullets` - 添加項目符號

**用途**: 在標題後批量添加項目符號列表

**語法**:
```bash
python word_editor.py document.docx add-bullets "標題" "項目1" "項目2" "項目3" ...
```

**使用時機**:
- 「在...下面列出以下項目」
- 添加清單、列表

**範例**:
```bash
python word_editor.py doc.docx add-bullets "主要功能" "功能A" "功能B" "功能C"
```

---

### Word 編輯器使用策略

#### 決策樹

```
用戶請求 → 是否需要查看結構？
           ↓
           是 → 執行 list
           ↓
           否 → 判斷操作類型
                ↓
                替換文字 → replace
                添加內容 → insert-after-heading / add-after
                刪除內容 → delete
                添加列表 → add-bullets
```

---

## PowerPoint 編輯器詳解

### 基本語法

```bash
python ppt_editor.py <檔案路徑> <命令> [參數] [選項]
```

### 命令清單

#### 1. `list` - 查看投影片結構

**用途**: 列出所有投影片的標題和內容摘要

**語法**:
```bash
python ppt_editor.py presentation.pptx list
```

**輸出格式**:
```
[投影片 1] 封面標題
  • 內容預覽第一行...
  • 內容預覽第二行...

[投影片 2] 第二張標題
  • 內容...
```

**何時使用**:
- 用戶未指定投影片編號
- 需要了解簡報結構
- 尋找特定內容所在位置

---

#### 2. `replace` - 替換文字

**用途**: 替換投影片中的文字（全部或指定投影片）

**語法**:
```bash
python ppt_editor.py presentation.pptx replace "舊文字" "新文字" [--slide N]
```

**參數**:
- `old`: 要替換的文字
- `new`: 新文字
- `--slide N`: (可選) 指定投影片編號

**範例**:
```bash
# 替換整份簡報
python ppt_editor.py slides.pptx replace "2024" "2025"

# 只替換第 3 張投影片
python ppt_editor.py slides.pptx replace "舊講師" "新講師" --slide 3
```

**注意**: 會搜尋文字框和表格

---

#### 3. `update-title` - 更新投影片標題

**用途**: 修改指定投影片的標題

**語法**:
```bash
python ppt_editor.py presentation.pptx update-title <投影片編號> "新標題"
```

**使用時機**:
- 「把第 N 張投影片的標題改成...」
- 更新章節標題

**範例**:
```bash
python ppt_editor.py slides.pptx update-title 3 "系統架構介紹"
```

---

#### 4. `add-slide` - 新增投影片

**用途**: 在簡報最後新增投影片

**語法**:
```bash
python ppt_editor.py presentation.pptx add-slide "標題" [--layout N]
```

**參數**:
- `title`: 新投影片的標題
- `--layout N`: (可選) 版面配置索引，預設為 1

**使用時機**:
- 「在最後加一張...」
- 新增感謝頁、Q&A 頁

**範例**:
```bash
python ppt_editor.py slides.pptx add-slide "Q&A"
python ppt_editor.py slides.pptx add-slide "感謝聆聽"
```

---

#### 5. `delete-slide` - 刪除投影片

**用途**: 刪除指定投影片

**語法**:
```bash
python ppt_editor.py presentation.pptx delete-slide <投影片編號>
```

**使用時機**:
- 「刪除第 N 張投影片」
- 移除不需要的頁面

**範例**:
```bash
python ppt_editor.py slides.pptx delete-slide 5
```

**警告**: 不可逆，建議先備份

---

#### 6. `add-text` - 添加文字

**用途**: 在投影片中添加文字內容

**語法**:
```bash
python ppt_editor.py presentation.pptx add-text <投影片編號> "文字內容"
```

**範例**:
```bash
python ppt_editor.py slides.pptx add-text 2 "補充說明：此為範例"
```

---

#### 7. `set-font` - 設定字體

**用途**: 修改投影片的字體和大小

**語法**:
```bash
python ppt_editor.py presentation.pptx set-font <投影片編號> "字體名稱" [--size N]
```

**範例**:
```bash
# 只改字體
python ppt_editor.py slides.pptx set-font 1 "微軟正黑體"

# 改字體和大小
python ppt_editor.py slides.pptx set-font 1 "微軟正黑體" --size 24
```

---

#### 8. `info` - 查看投影片詳細資訊

**用途**: 顯示指定投影片的完整內容

**語法**:
```bash
python ppt_editor.py presentation.pptx info <投影片編號>
```

**使用時機**:
- 需要查看特定投影片的詳細內容
- 確認修改前的狀態

---

### PowerPoint 編輯器使用策略

#### 投影片編號注意事項

- **從 1 開始**: 投影片編號從 1 開始（不是 0）
- **動態變化**: 刪除投影片後，後續投影片編號會前移
- **先查看**: 不確定編號時，先執行 `list`

---

## 自然語言到命令的轉換

### 關鍵詞映射表

#### Word 文檔

| 用戶意圖 | 關鍵詞 | 命令 | 範例 |
|---------|--------|------|------|
| 查看結構 | 「顯示」「列出」「有哪些」 | `list` | "顯示文檔結構" → `list` |
| 替換文字 | 「把...改成」「替換」「更新為」 | `replace` | "把 A 改成 B" → `replace "A" "B"` |
| 添加內容 | 「在...後面加」「插入」 | `insert-after-heading` | "在標題後加段落" → `insert-after-heading` |
| 刪除內容 | 「刪除」「移除」 | `delete` | "刪除某段" → `delete` |
| 添加列表 | 「列出」「項目」 | `add-bullets` | "加項目符號" → `add-bullets` |

#### PowerPoint

| 用戶意圖 | 關鍵詞 | 命令 | 範例 |
|---------|--------|------|------|
| 查看投影片 | 「顯示」「有哪些投影片」 | `list` | "顯示所有投影片" → `list` |
| 替換文字 | 「把...改成」「替換」 | `replace` | "把講師名改成..." → `replace` |
| 改標題 | 「標題改成」「更新標題」 | `update-title` | "第3張標題改成..." → `update-title 3` |
| 新增投影片 | 「新增」「加一張」 | `add-slide` | "最後加一張" → `add-slide` |
| 刪除投影片 | 「刪除」「移除」 | `delete-slide` | "刪除第5張" → `delete-slide 5` |
| 改字體 | 「字體」「字型」 | `set-font` | "改成正黑體" → `set-font` |

### 轉換流程

```
1. 解析用戶意圖
   ↓
2. 識別目標對象（Word/PPT，哪個檔案）
   ↓
3. 提取關鍵參數（文字、編號、位置）
   ↓
4. 映射到對應命令
   ↓
5. 組裝完整命令
   ↓
6. 執行並回報結果
```

### 範例轉換

#### 範例 1
```
用戶: "把簡報中所有的 2024 改成 2025"

分析:
- 對象: PowerPoint
- 意圖: 替換文字
- 參數: "2024" → "2025"
- 範圍: 所有投影片

命令:
python ppt_editor.py presentation.pptx replace "2024" "2025"
```

#### 範例 2
```
用戶: "將第 3 張投影片的標題改成『系統架構』"

分析:
- 對象: PowerPoint
- 意圖: 更新標題
- 參數: 投影片 3, 新標題 "系統架構"

命令:
python ppt_editor.py presentation.pptx update-title 3 "系統架構"
```

#### 範例 3
```
用戶: "在『課程目標』這個標題後面，加上『本課程旨在...』這段話"

分析:
- 對象: Word
- 意圖: 在標題後插入
- 參數: 標題 "課程目標", 內容 "本課程旨在..."

命令:
python word_editor.py document.docx insert-after-heading "課程目標" "本課程旨在..."
```

---

## 最佳實踐

### 1. 工作流程建議

#### 標準流程
```
1. 理解需求
   - 明確用戶要操作的檔案
   - 確認具體要做什麼

2. 查看結構（如有需要）
   - 執行 list 命令
   - 確認目標位置

3. 執行操作
   - 選擇正確命令
   - 提供準確參數

4. 確認結果
   - 閱讀命令輸出
   - 確認操作成功

5. 回報用戶
   - 簡潔說明完成的操作
   - 提供關鍵資訊（如替換了幾處）
```

### 2. 安全考量

- **建議備份**: 對於刪除、大量替換操作，建議用戶先備份
- **使用 --output**: 重要文件使用 `--output` 另存新檔
- **確認再執行**: 刪除操作前確認用戶意圖

### 3. 錯誤預防

- **檔案路徑**: 確認檔案存在且路徑正確
- **編號範圍**: 投影片/段落編號在有效範圍內
- **文字大小寫**: `replace` 區分大小寫，提醒用戶
- **特殊字元**: 含空格或特殊字元的文字用引號包起來

### 4. 效率提升

- **批次操作**: 多個相似操作可以連續執行
- **優先 list**: 不確定時先查看結構
- **善用指定範圍**: 用 `--slide` 限縮範圍提高精確度

---

## 錯誤處理

### 常見錯誤與處理

#### 1. 檔案不存在
```
錯誤訊息: "無法開啟檔案"
處理: 確認檔案路徑，詢問用戶正確位置
```

#### 2. 找不到目標文字
```
錯誤訊息: "找不到「XXX」"
處理: 
- 確認文字拼寫
- 建議執行 list 查看
- 檢查大小寫
```

#### 3. 投影片編號超出範圍
```
錯誤訊息: "投影片編號 X 不存在"
處理:
- 執行 list 查看總數
- 提供正確範圍
```

#### 4. 沒有標題框
```
錯誤訊息: "投影片沒有標題框"
處理:
- 說明該投影片無標題
- 建議使用其他命令
```

### 錯誤回應範本

```
❌ 操作失敗: [錯誤原因]

建議:
1. [建議動作 1]
2. [建議動作 2]

需要我幫您 [替代方案] 嗎？
```

---

## 完整工作流程範例

### 場景 1: 更新培訓簡報的年份和講師

```
用戶: "請把培訓簡報中的 2024 改成 2025，講師從 John 改成 Sarah"

Agent 操作流程:

1. 執行第一個替換
Command: python ppt_editor.py training.pptx replace "2024" "2025"
Output: ✓ 替換了 15 處「2024」→「2025」

2. 執行第二個替換
Command: python ppt_editor.py training.pptx replace "John" "Sarah"
Output: ✓ 替換了 8 處「John」→「Sarah」

3. 回報用戶
"✅ 已完成更新：
- 將所有「2024」改為「2025」（15 處）
- 將所有「John」改為「Sarah」（8 處）

簡報已儲存！"
```

### 場景 2: Word 文檔添加新章節

```
用戶: "在『系統概述』後面加一個二級標題『主要特色』，然後列出三個特色"

Agent 操作流程:

1. 插入標題
Command: python word_editor.py doc.docx insert-after-heading "系統概述" "主要特色" --is-heading --heading-level 2
Output: ✓ 已在標題「系統概述」後插入內容

2. 添加項目符號
Command: python word_editor.py doc.docx add-bullets "主要特色" "即時監控" "智能保護" "自動化測試"
Output: ✓ 已在「主要特色」後添加 3 個項目

3. 回報用戶
"✅ 已完成：
- 在「系統概述」後添加了二級標題「主要特色」
- 添加了 3 個項目符號列表

文檔已更新！"
```

### 場景 3: 複雜的 PPT 重組

```
用戶: "簡報太長了，幫我刪除第 5 和第 7 張，然後在最後加一張 Q&A"

Agent 操作流程:

1. 先查看結構（確認）
Command: python ppt_editor.py long.pptx list
Output: [顯示所有投影片]

2. 刪除第 5 張
Command: python ppt_editor.py long.pptx delete-slide 5
Output: ✓ 已刪除投影片 5: [標題]

3. 刪除第 6 張（原本的第 7 張）
注意: 刪除後編號會改變！
Command: python ppt_editor.py long.pptx delete-slide 6
Output: ✓ 已刪除投影片 6: [標題]

4. 新增 Q&A 投影片
Command: python ppt_editor.py long.pptx add-slide "Q&A"
Output: ✓ 已新增投影片 [N]: Q&A

5. 回報用戶
"✅ 已完成簡報調整：
- 刪除了原第 5 張投影片
- 刪除了原第 7 張投影片
- 在最後新增了 Q&A 投影片

現在簡報共有 [N] 張投影片。"
```

---

## 進階技巧

### 1. 批次操作腳本

當需要重複操作多張投影片時，可以連續執行命令：

```bash
# 為每張投影片設定統一字體
python ppt_editor.py slides.pptx set-font 1 "微軟正黑體" --size 20
python ppt_editor.py slides.pptx set-font 2 "微軟正黑體" --size 20
python ppt_editor.py slides.pptx set-font 3 "微軟正黑體" --size 20
```

### 2. 使用 --output 保護原檔

```bash
python word_editor.py original.docx replace "test" "production" --output final.docx
```

### 3. 組合多個工具

可以在同一個工作流程中使用 Word 和 PPT 編輯器：

```
1. 更新 Word 文檔的內容
2. 根據更新後的內容修改對應的 PPT
3. 確保兩者一致
```

---

## 快速參考卡

### Word 編輯器命令

```
list                              # 查看結構
replace "A" "B"                   # 替換文字
add-after "搜尋" "內容"            # 段落後添加
insert-after-heading "標題" "內容" # 標題後插入
delete "搜尋"                      # 刪除段落
add-bullets "標題" "1" "2" "3"    # 添加列表
```

### PPT 編輯器命令

```
list                              # 查看投影片
replace "A" "B" [--slide N]       # 替換文字
update-title N "新標題"             # 更新標題
add-slide "標題"                   # 新增投影片
delete-slide N                    # 刪除投影片
add-text N "內容"                  # 添加文字
set-font N "字體" [--size N]       # 設定字體
info N                            # 查看詳情
```

### 通用選項

```
--output file.docx/pptx           # 另存新檔
```

---

## 總結

### 關鍵要點

1. ✅ **先理解再操作**: 不確定時使用 `list` 查看
2. ✅ **精確指定**: 提供明確的參數避免錯誤
3. ✅ **安全第一**: 重要操作建議備份或用 `--output`
4. ✅ **驗證結果**: 讀取命令輸出確認成功
5. ✅ **清楚溝通**: 向用戶簡潔報告操作結果

### Agent 自檢清單

執行命令前檢查:
- [ ] 檔案路徑正確？
- [ ] 命令和參數正確？
- [ ] 需要備份嗎？
- [ ] 引號使用正確？
- [ ] 編號在有效範圍內？

執行命令後檢查:
- [ ] 命令成功執行？
- [ ] 輸出訊息顯示預期結果？
- [ ] 需要執行後續命令？
- [ ] 向用戶報告結果？

---

**版本歷史**
- v1.0 (2025-12-01): 初版發布，包含 Word 和 PPT 編輯器完整說明
