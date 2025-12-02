"""Office Document Editor Tools

A suite of command-line tools for editing Microsoft Office documents.
"""

__version__ = "1.3.0"
__author__ = "Development Team"

from .word_editor import WordEditor
from .ppt_editor import PPTEditor
from .excel_editor import ExcelEditor
from .batch_processor import BatchProcessor
from . import llm_api

__all__ = ['WordEditor', 'PPTEditor', 'ExcelEditor', 'BatchProcessor', 'llm_api']
