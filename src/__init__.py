"""Office Document Editor Tools

A suite of command-line tools for editing Microsoft Office documents.
"""

__version__ = "1.1.0"
__author__ = "Development Team"

from .word_editor import WordEditor
from .ppt_editor import PPTEditor
from .excel_editor import ExcelEditor

__all__ = ['WordEditor', 'PPTEditor', 'ExcelEditor']
