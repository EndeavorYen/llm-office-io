"""
Unit tests for WordEditor
"""

import unittest
import os
import sys
from pathlib import Path

# Add src to path
sys.path.insert(0, str(Path(__file__).parent.parent / 'src'))

try:
    from src.word_editor import WordEditor
    from src.constants import WORD_EXTENSION
except ImportError:
    # Try direct import if running from tests directory
    import word_editor
    import constants
    WordEditor = word_editor.WordEditor
    WORD_EXTENSION = constants.WORD_EXTENSION


class TestWordEditorInitialization(unittest.TestCase):
    """測試 WordEditor 初始化"""
    
    def test_init_with_nonexistent_file(self):
        """測試不存在的檔案"""
        with self.assertRaises(FileNotFoundError):
            WordEditor("nonexistent.docx")
    
    def test_init_with_invalid_extension(self):
        """測試無效的檔案格式"""
        with self.assertRaises(ValueError):
            WordEditor("file.txt")
    
    def test_word_extension_constant(self):
        """測試常量定義"""
        self.assertEqual(WORD_EXTENSION, '.docx')


class TestWordEditorReplaceText(unittest.TestCase):
    """測試文字替換功能"""
    
    def test_replace_with_empty_old_text(self):
        """測試空字串替換"""
        # 這個測試需要一個真實的docx檔案
        # 暫時跳過，需要測試fixtures
        pass
    
    def test_replace_count_parameter(self):
        """測試替換次數參數"""
        # 需要測試fixtures
        pass


class TestInputValidation(unittest.TestCase):
    """測試輸入驗證"""
    
    def test_empty_search_text_in_delete(self):
        """測試空搜尋文字"""
        # 需要測試fixtures
        pass


def suite():
    """建立測試套件"""
    test_suite = unittest.TestSuite()
    test_suite.addTest(unittest.makeSuite(TestWordEditorInitialization))
    test_suite.addTest(unittest.makeSuite(TestWordEditorReplaceText))
    test_suite.addTest(unittest.makeSuite(TestInputValidation))
    return test_suite


if __name__ == '__main__':
    runner = unittest.TextTestRunner(verbosity=2)
    runner.run(suite())
