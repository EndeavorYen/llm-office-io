# PowerPoint Editor 使用指南

## 📖 工具介紹

`ppt_editor.py` 是一個強大的 PowerPoint 互動式編輯工具，讓您可以透過簡單的命令列指令修改 PPT 檔案。

## 🎯 主要功能

### 1. 查看投影片結構
列出所有投影片的標題和內容摘要

```bash
python ppt_editor.py presentation.pptx list
```

### 2. 替換文字
在整份簡報或特定投影片中替換文字

```bash
# 替換所有投影片的文字
python ppt_editor.py presentation.pptx replace "舊文字" "新文字"

# 只替換第 3 張投影片的文字
python ppt_editor.py presentation.pptx replace "舊文字" "新文字" --slide 3
```

### 3. 更新投影片標題
修改指定投影片的標題

```bash
python ppt_editor.py presentation.pptx update-title 3 "新的標題"
```

### 4. 新增投影片
在簡報最後新增一張投影片

```bash
python ppt_editor.py presentation.pptx add-slide "新投影片標題"
```

### 5. 刪除投影片
刪除指定的投影片

```bash
python ppt_editor.py presentation.pptx delete-slide 5
```

### 6. 添加文字
在指定投影片中添加文字內容

```bash
python ppt_editor.py presentation.pptx add-text 2 "新增的內容"
```

### 7. 設定字體
修改指定投影片的字體

```bash
# 只改字體
python ppt_editor.py presentation.pptx set-font 1 "微軟正黑體"

# 改字體和大小
python ppt_editor.py presentation.pptx set-font 1 "微軟正黑體" --size 18
```

### 8. 查看投影片詳細資訊
顯示指定投影片的完整資訊

```bash
python ppt_editor.py presentation.pptx info 3
```

## 💾 輸出設定

預設會覆蓋原檔案，如果要另存新檔：

```bash
python ppt_editor.py input.pptx replace "舊" "新" --output output.pptx
```

## 📝 使用範例

### 範例 1：批量替換講師名稱
```bash
python ppt_editor.py training.pptx replace "John" "Sarah"
```

### 範例 2：更新第一張投影片標題
```bash
python ppt_editor.py training.pptx update-title 1 "2025 年度培訓課程"
```

### 範例 3：在最後新增感謝頁
```bash
python ppt_editor.py training.pptx add-slide "感謝聆聽"
```

### 範例 4：設定所有投影片統一字體
```bash
# 需要逐張設定（可以寫 shell 腳本批次執行）
python ppt_editor.py training.pptx set-font 1 "微軟正黑體" --size 20
python ppt_editor.py training.pptx set-font 2 "微軟正黑體" --size 20
# ... 以此類推
```

## 🔧 技術細節

- 支援 .pptx 格式（PowerPoint 2007+）
- 可處理文字框、表格中的文字
- 保留原有格式和樣式（顏色、對齊等）
- 支援中英文字型設定

## ⚠️ 注意事項

1. **備份重要**：修改前建議先備份原檔案
2. **複雜圖形**：無法編輯 SmartArt、圖表等複雜物件
3. **動畫效果**：不會影響現有動畫設定
4. **版面配置**：新增投影片使用預設版面配置

## 🚀 自然語言使用方式

您可以直接告訴我要做什麼，例如：

- 「把所有的『2024』改成『2025』」
- 「將第 5 張投影片的標題改成『總結』」
- 「在簡報最後加一張『Q&A』投影片」
- 「刪除第 3 張投影片」
- 「顯示所有投影片的標題」

我會自動執行對應的命令！✨
