"""
Unit tests for PPTEditor
"""

import unittest
import os
import sys
from pathlib import Path

# Add src to path
sys.path.insert(0, str(Path(__file__).parent.parent / 'src'))

try:
    from src.ppt_editor import PPTEditor
    from src.constants import PPT_EXTENSION, DEFAULT_LAYOUT_INDEX
except ImportError:
    # Try direct import if running from tests directory
    import ppt_editor
    import constants
    PPTEditor = ppt_editor.PPTEditor
    PPT_EXTENSION = constants.PPT_EXTENSION
    DEFAULT_LAYOUT_INDEX = constants.DEFAULT_LAYOUT_INDEX


class TestPPTEditorInitialization(unittest.TestCase):
    """測試 PPTEditor 初始化"""
    
    def test_init_with_nonexistent_file(self):
        """測試不存在的檔案"""
        with self.assertRaises(FileNotFoundError):
            PPTEditor("nonexistent.pptx")
    
    def test_init_with_invalid_extension(self):
        """測試無效的檔案格式"""
        with self.assertRaises(ValueError):
            PPTEditor("file.txt")
    
    def test_ppt_extension_constant(self):
        """測試常量定義"""
        self.assertEqual(PPT_EXTENSION, '.pptx')
    
    def test_default_layout_index(self):
        """測試預設版面配置索引"""
        self.assertEqual(DEFAULT_LAYOUT_INDEX, 1)


class TestPPTEditorReplaceText(unittest.TestCase):
    """測試文字替換功能"""
    
    def test_replace_with_empty_old_text(self):
        """測試空字串替換"""
        # 需要測試fixtures
        pass


class TestSlideValidation(unittest.TestCase):
    """測試投影片編號驗證"""
    
    def test_invalid_slide_number(self):
        """測試無效的投影片編號"""
        # 需要測試fixtures
        pass


def suite():
    """建立測試套件"""
    test_suite = unittest.TestSuite()
    test_suite.addTest(unittest.makeSuite(TestPPTEditorInitialization))
    test_suite.addTest(unittest.makeSuite(TestPPTEditorReplaceText))
    test_suite.addTest(unittest.makeSuite(TestSlideValidation))
    return test_suite


if __name__ == '__main__':
    runner = unittest.TextTestRunner(verbosity=2)
    runner.run(suite())
