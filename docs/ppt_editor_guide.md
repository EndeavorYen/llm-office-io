# PowerPoint Editor 完整使用指南

## 📖 簡介

PowerPoint Editor 提供 12 個強大功能，讓您輕鬆自動化簡報的編輯操作。

---

## 🚀 快速開始

```python
from src.ppt_editor import PPTEditor

# 開啟簡報
editor = PPTEditor("presentation.pptx")

# 執行操作
editor.replace_text("舊文字", "新文字")
editor.save("output.pptx")
```

---

## 📋 功能列表

### 1. 文字替換 `replace_text()`

```python
# 替換所有投影片的文字
count = editor.replace_text("2024", "2025")

# 只替換特定投影片
count = editor.replace_text("Draft", "Final", slide_number=3)
```

---

### 2. 插入圖片 `add_image()` 🆕

```python
# 在投影片 2 插入圖片
editor.add_image(
    slide_number=2,
    image_path="chart.png",
    left_cm=5.0,      # 左邊距 5cm
    top_cm=8.0,       # 上邊距 8cm
    width_cm=15.0     # 寬度 15cm
)

# 插入公司標誌（右上角）
editor.add_image(
    slide_number=1,
    image_path="logo.png",
    left_cm=22.0,
    top_cm=1.0,
    width_cm=3.0
)
```

---

### 3. 添加文字方塊 `add_textbox()` 🆕

```python
# 在投影片 3 添加文字方塊
editor.add_textbox(
    slide_number=3,
    text="重要提示：請注意時程安排",
    left_cm=2.0,
    top_cm=12.0,
    width_cm=20.0,
    height_cm=3.0,
    font_size=24
)
```

---

### 4. 添加形狀 `add_shape()` 🆕

```python
# 添加矩形
editor.add_shape(
    slide_number=4,
    shape_type='rectangle',
    left_cm=5.0,
    top_cm=10.0,
    width_cm=15.0,
    height_cm=5.0,
    fill_color=(255, 200, 100)  # 橙色 RGB
)

# 添加橢圓
editor.add_shape(
    slide_number=5,
    shape_type='oval',
    left_cm=10.0,
    top_cm=8.0,
    width_cm=8.0,
    height_cm=8.0,
    fill_color=(100, 150, 255)  # 藍色
)

# 添加圓角矩形
editor.add_shape(
    slide_number=6,
    shape_type='rounded_rectangle',
    fill_color=(0, 200, 0)  # 綠色
)
```

**支援的形狀**: `'rectangle'`, `'oval'`, `'rounded_rectangle'`

---

### 5. 複製投影片 `duplicate_slide()` 🆕

```python
# 複製投影片 3
editor.duplicate_slide(slide_number=3)
# 新投影片會添加到簡報最後
```

---

### 6. 設定背景顏色 `set_background_color()` 🆕

```python
# 設定投影片 1 背景為白色
editor.set_background_color(
    slide_number=1,
    color=(255, 255, 255)  # RGB
)

# 設定淺藍色背景
editor.set_background_color(
    slide_number=2,
    color=(230, 240, 255)
)

# 常用顏色
# 白色: (255, 255, 255)
# 黑色: (0, 0, 0)
# 淺灰: (240, 240, 240)
# 淺藍: (230, 240, 255)
# 淺綠: (230, 255, 230)
```

---

### 7. 更新投影片標題 `update_slide_title()`

```python
# 更新第 2 張投影片的標題
editor.update_slide_title(
    slide_number=2,
    new_title="新的標題文字"
)
```

---

### 8. 新增投影片 `add_slide()`

```python
# 新增投影片（使用預設版面）
editor.add_slide("新投影片標題")

# 使用特定版面配置
editor.add_slide("標題投影片", layout_index=0)
```

---

### 9. 刪除投影片 `delete_slide()`

```python
# 刪除第 5 張投影片
editor.delete_slide(slide_number=5)
```

---

### 10. 列出所有投影片 `list_slides()`

```python
# 顯示所有投影片的標題和內容預覽
editor.list_slides()
```

輸出範例:
```
=== 簡報結構 (共 5 張投影片) ===

📊 投影片 1: 年度報告
  內容: 2024年度業績總結...

📊 投影片 2: 財務摘要
  內容: 營收成長 15%...
```

---

### 11. 查看單張投影片 `view_slide()`

```python
# 查看第 3 張投影片的詳細內容
editor.view_slide(slide_number=3)
```

---

### 12. 儲存簡報 `save()`

```python
# 覆蓋原檔案
editor.save()

# 另存新檔
editor.save("new_presentation.pptx")
```

---

## 💡 實用範例

### 範例 1: 品牌簡報製作

```python
editor = PPTEditor("template.pptx")

# 所有投影片加上公司標誌
for i in range(1, len(editor.prs.slides) + 1):
    editor.add_image(
        slide_number=i,
        image_path="company_logo.png",
        left_cm=22.0,
        top_cm=1.0,
        width_cm=3.0
    )

# 設定標題投影片背景
editor.set_background_color(1, (0, 51, 102))  # 深藍色

# 更新年份
editor.replace_text("2024", "2025")

editor.save("branded_presentation.pptx")
```

---

### 範例 2: 資料視覺化簡報

```python
editor = PPTEditor("data_report.pptx")

# 插入圖表圖片
editor.add_image(
    slide_number=3,
    image_path="sales_chart.png",
    left_cm=3.0,
    top_cm=5.0,
    width_cm=20.0
)

# 添加說明文字
editor.add_textbox(
    slide_number=3,
    text="營收成長趨勢（2024 Q1-Q4）",
    left_cm=3.0,
    top_cm=4.0,
    width_cm=20.0,
    height_cm=1.5,
    font_size=18
)

# 添加重點標記
editor.add_shape(
    slide_number=3,
    shape_type='oval',
    left_cm=18.0,
    top_cm=10.0,
    width_cm=2.0,
    height_cm=2.0,
    fill_color=(255, 0, 0)  # 紅色圓圈標記
)

editor.save()
```

---

### 範例 3: 快速複製模板投影片

```python
editor = PPTEditor("quarterly_report.pptx")

# 假設投影片 5 是「月度摘要」模板
# 複製 3 次用於 Q2, Q3, Q4
for month in range(3):
    editor.duplicate_slide(slide_number=5)

# 更新每個月份的標題
editor.update_slide_title(6, "Q2 月度摘要")
editor.update_slide_title(7, "Q3 月度摘要")
editor.update_slide_title(8, "Q4 月度摘要")

editor.save()
```

---

### 範例 4: 添加視覺元素

```python
editor = PPTEditor("presentation.pptx")

# 在投影片 2 添加圖片和形狀組合
# 背景矩形
editor.add_shape(
    slide_number=2,
    shape_type='rounded_rectangle',
    left_cm=5.0,
    top_cm=8.0,
    width_cm=16.0,
    height_cm=8.0,
    fill_color=(240, 240, 240)  # 淺灰背景
)

# 產品圖片
editor.add_image(
    slide_number=2,
    image_path="product.png",
    left_cm=6.0,
    top_cm=9.0,
    width_cm=6.0
)

# 產品說明文字
editor.add_textbox(
    slide_number=2,
    text="全新產品特色：\n• 輕量設計\n• 高效能\n• 節能環保",
    left_cm=13.0,
    top_cm=9.0,
    width_cm=7.0,
    height_cm=6.0,
    font_size=14
)

editor.save()
```

---

## 📐 位置參考

PowerPoint 標準投影片尺寸（16:9）:
- **寬度**: 約 25.4 cm (10 inches)
- **高度**: 約 19.05 cm (7.5 inches)

常用位置:
- **左上角**: left_cm=1.0, top_cm=1.0
- **右上角**: left_cm=22.0, top_cm=1.0
- **中央**: left_cm=7.0, top_cm=7.0
- **底部**: top_cm=16.0

---

## ⚠️ 注意事項

1. **檔案格式**: 僅支援 `.pptx` 格式
2. **投影片編號**: 從 1 開始（不是 0）
3. **RGB 顏色**: 範圍 0-255
4. **位置單位**: 使用公分（cm）

---

## 🎯 最佳實踐

1. **視覺一致性**: 使用相同的顏色和字體大小
2. **複製模板**: 使用 `duplicate_slide()` 保持格式一致
3. **測試位置**: 先在單張投影片測試位置參數
4. **批次處理**: 使用迴圈處理多張投影片

---

更多範例請參考 [examples/](../examples/) 目錄。
