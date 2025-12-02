# PPT Editor - Advanced Features Extension
# This file contains the advanced features to be added to ppt_editor.py

def add_image(self, slide_number: int, image_path: str, 
             left_cm: float = 2.0, top_cm: float = 5.0, 
             width_cm: float = 10.0) -> bool:
    """在投影片中插入圖片
    
    Args:
        slide_number: 投影片編號（從 1 開始）
        image_path: 圖片檔案路徑
        left_cm: 左邊距（公分）
        top_cm: 上邊距（公分）
        width_cm: 圖片寬度（公分）
        
    Returns:
        bool: 是否成功插入
    """
    if not self._validate_slide_number(slide_number):
        return False
    
    if not os.path.exists(image_path):
        print(f"{ERROR_SYMBOL} 圖片檔案不存在: {image_path}")
        return False
    
    try:
        slide = self.prs.slides[slide_number - 1]
        slide.shapes.add_picture(
            image_path,
            Cm(left_cm),
            Cm(top_cm),
            width=Cm(width_cm)
        )
        print(f"{SUCCESS_SYMBOL} 已在投影片 {slide_number} 插入圖片")
        return True
    except Exception as e:
        print(f"{ERROR_SYMBOL} 插入圖片失敗: {e}")
        return False

def add_textbox(self, slide_number: int, text: str,
               left_cm: float = 2.0, top_cm: float = 2.0,
               width_cm: float = 15.0, height_cm: float = 3.0,
               font_size: int = 18) -> bool:
    """在投影片中添加文字方塊
    
    Args:
        slide_number: 投影片編號（從 1 開始）
        text: 文字內容
        left_cm: 左邊距（公分）
        top_cm: 上邊距（公分）
        width_cm: 寬度（公分）
        height_cm: 高度（公分）
        font_size: 字體大小
        
    Returns:
        bool: 是否成功添加
    """
    if not self._validate_slide_number(slide_number):
        return False
    
    try:
        slide = self.prs.slides[slide_number - 1]
        textbox = slide.shapes.add_textbox(
            Cm(left_cm),
            Cm(top_cm),
            Cm(width_cm),
            Cm(height_cm)
        )
        text_frame = textbox.text_frame
        text_frame.text = text
        
        # 設定字體大小
        for paragraph in text_frame.paragraphs:
            for run in paragraph.runs:
                run.font.size = Pt(font_size)
        
        print(f"{SUCCESS_SYMBOL} 已在投影片 {slide_number} 添加文字方塊")
        return True
    except Exception as e:
        print(f"{ERROR_SYMBOL} 添加文字方塊失敗: {e}")
        return False

def add_shape(self, slide_number: int, shape_type: str = 'rectangle',
             left_cm: float = 5.0, top_cm: float = 5.0,
             width_cm: float = 10.0, height_cm: float = 5.0,
             fill_color: Optional[tuple] = None) -> bool:
    """在投影片中添加形狀
    
    Args:
        slide_number: 投影片編號（從 1 開始）
        shape_type: 形狀類型 ('rectangle', 'oval', 'rounded_rectangle')
        left_cm: 左邊距（公分）
        top_cm: 上邊距（公分）
        width_cm: 寬度（公分）
        height_cm: 高度（公分）
        fill_color: 填充顏色 RGB (r, g, b)，例如 (255, 0, 0) 為紅色
        
    Returns:
        bool: 是否成功添加
    """
    if not self._validate_slide_number(slide_number):
        return False
    
    shape_map = {
        'rectangle': MSO_SHAPE.RECTANGLE,
        'oval': MSO_SHAPE.OVAL,
        'rounded_rectangle': MSO_SHAPE.ROUNDED_RECTANGLE,
    }
    
    if shape_type not in shape_map:
        print(f"{ERROR_SYMBOL} 不支援的形狀類型: {shape_type}")
        print(f"  支援的類型: {', '.join(shape_map.keys())}")
        return False
    
    try:
        slide = self.prs.slides[slide_number - 1]
        shape = slide.shapes.add_shape(
            shape_map[shape_type],
            Cm(left_cm),
            Cm(top_cm),
            Cm(width_cm),
            Cm(height_cm)
        )
        
        # 設定填充顏色
        if fill_color:
            fill = shape.fill
            fill.solid()
            fill.fore_color.rgb = RGBColor(*fill_color)
        
        print(f"{SUCCESS_SYMBOL} 已在投影片 {slide_number} 添加 {shape_type}")
        return True
    except Exception as e:
        print(f"{ERROR_SYMBOL} 添加形狀失敗: {e}")
        return False

def duplicate_slide(self, slide_number: int) -> bool:
    """複製投影片
    
    Args:
        slide_number: 要複製的投影片編號（從 1 開始）
        
    Returns:
        bool: 是否成功複製
    """
    if not self._validate_slide_number(slide_number):
        return False
    
    try:
        import copy
        source_slide = self.prs.slides[slide_number - 1]
        
        # 使用相同的版面配置創建新投影片
        blank_slide_layout = source_slide.slide_layout
        new_slide = self.prs.slides.add_slide(blank_slide_layout)
        
        # 複製所有形狀
        for shape in source_slide.shapes:
            el = shape.element
            newel = copy.deepcopy(el)
            new_slide.shapes._spTree.insert_element_before(newel, 'p:extLst')
        
        new_number = len(self.prs.slides)
        print(f"{SUCCESS_SYMBOL} 已複製投影片 {slide_number} → 投影片 {new_number}")
        return True
    except Exception as e:
        print(f"{ERROR_SYMBOL} 複製投影片失敗: {e}")
        return False

def set_background_color(self, slide_number: int, color: tuple) -> bool:
    """設定投影片背景顏色
    
    Args:
        slide_number: 投影片編號（從 1 開始）
        color: RGB 顏色 (r, g, b)，例如 (255, 255, 255) 為白色
        
    Returns:
        bool: 是否成功設定
    """
    if not self._validate_slide_number(slide_number):
        return False
    
    try:
        slide = self.prs.slides[slide_number - 1]
        background = slide.background
        fill = background.fill
        fill.solid()
        fill.fore_color.rgb = RGBColor(*color)
        
        print(f"{SUCCESS_SYMBOL} 已設定投影片 {slide_number} 背景顏色為 RGB{color}")
        return True
    except Exception as e:
        print(f"{ERROR_SYMBOL} 設定背景顏色失敗: {e}")
        return False
