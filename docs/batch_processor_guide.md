# Batch Processor 使用指南

## 📖 簡介

`batch_processor.py` 讓您可以一次處理多個 Office 文檔，極大提升工作效率！

## 🚀 快速開始

### 基本語法

```bash
python src/batch_processor.py <檔案模式> <命令> [參數] [選項]
```

## 💡 常用範例

### 範例 1: 批次替換 Word 文檔

```bash
# 替換當前目錄所有 Word 文檔
python src/batch_processor.py "*.docx" replace "2024" "2025"

# 替換特定目錄的檔案
python src/batch_processor.py "reports/*.docx" replace "舊版本" "新版本"
```

### 範例 2: 遞迴處理所有子目錄

```bash
# 處理所有子目錄的 Excel 檔案
python src/batch_processor.py "*.xlsx" replace "Draft" "Final" --recursive
```

### 範例 3: 輸出到指定目錄

```bash
# 將處理後的檔案儲存到 output 目錄
python src/batch_processor.py "*.pptx" replace "Q3" "Q4" --output output/
```

### 範例 4: 備份原檔案

```bash
# 處理前自動備份（會創建 .bak 檔案）
python src/batch_processor.py "*.docx" replace "A" "B" --backup
```

### 範例 5: 完整範例

```bash
# 遞迴處理、輸出到新目錄、備份原檔案
python src/batch_processor.py "**/*.xlsx" replace "舊" "新" -r -o processed/ -b
```

## 📋 支援的命令

### replace
批次替換文字

```bash
python src/batch_processor.py "*.docx" replace "舊文字" "新文字"
```

### delete
刪除包含特定文字的段落（僅 Word）

```bash
python src/batch_processor.py "*.docx" delete "待刪除的內容"
```

## ⚙️ 選項說明

| 選項 | 簡寫 | 說明 |
|------|------|------|
| --recursive | -r | 遞迴搜尋所有子目錄 |
| --output DIR | -o DIR | 將結果輸出到指定目錄 |
| --backup | -b | 處理前備份原檔案（.bak） |

## 📊 輸出說明

處理時會顯示：
- 找到的檔案列表
- 處理進度（如果安裝了 tqdm）
- 每個檔案的處理結果
- 最終統計資訊

範例輸出：
```
找到 15 個檔案
  - report1.docx
  - report2.docx
  ...

執行命令: replace 2024 2025

處理檔案: 100%|██████████| 15/15 [00:03<00:00,  4.2it/s]

==================================================
處理完成!
  總數: 15
  ✓ 成功: 14
  ✗ 失敗: 1
==================================================
```

## ⚠️ 注意事項

1. **檔案模式** - 使用引號包住檔案模式，如 `"*.docx"`
2. **備份建議** - 處理重要檔案時建議使用 `--backup`
3. **測試先行** - 先用少量檔案測試
4. **進度顯示** - 安裝 tqdm 可顯示進度條: `pip install tqdm`

## 🎯 實際應用場景

### 場景 1: 年度文檔更新
```bash
# 更新所有報告的年份
python src/batch_processor.py "reports/**/*.docx" replace "2024" "2025" -r -b
```

### 場景 2: 品牌名稱變更
```bash
# 更新所有簡報的公司名稱
python src/batch_processor.py "presentations/*.pptx" replace "舊公司名" "新公司名" -o updated/
```

### 場景 3: 統一術語
```bash
# 統一所有文檔的專業術語
python src/batch_processor.py "docs/*.docx" replace "客戶" "客戶" --backup
```

---

**提示**: 結合使用 `--recursive`, `--output`, `--backup` 可以安全高效地批次處理大量檔案！
