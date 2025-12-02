#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Quick Test for LLM API
快速測試 LLM-friendly API 功能
"""

import sys
import os
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from src.llm_api import replace_text, execute_command, execute_json
from docx import Document


def test_basic_functions():
    """測試基本功能"""
    print("=== 測試 LLM API ===\n")
    
    # 創建測試文檔
    test_file = "test_llm_api.docx"
    doc = Document()
    doc.add_paragraph("這是測試文檔 2024")
    doc.save(test_file)
    
    # 1. 測試 replace_text
    print("1. 測試 replace_text()")
    result = replace_text(test_file, "2024", "2025")
    print(f"   結果: {result}")
    assert result["success"] == True
    assert result["file_type"] == "word"
    print("   ✓ 通過\n")
    
    # 2. 測試 execute_command
    print("2. 測試 execute_command()")
    result = execute_command(
        "replace_text",
        file_path=test_file,
        old_text="2025",
        new_text="2026"
    )
    print(f"   結果: {result}")
    assert result["success"] == True
    print("   ✓ 通過\n")
    
    # 3. 測試 JSON 模式
    print("3. 測試 execute_json()")
    json_input = {
        "command": "replace_text",
        "params": {
            "file_path": test_file,
            "old_text": "2026",
            "new_text": "2027"
        }
    }
    result_json = execute_json(json_input)
    print(f"   結果: {result_json[:100]}...")
    assert '"success": true' in result_json
    print("   ✓ 通過\n")
    
    # 4. 測試錯誤處理
    print("4. 測試錯誤處理")
    result = replace_text("nonexistent.docx", "A", "B")
    print(f"   結果: {result}")
    assert result["success"] == False
    assert result["error"] is not None
    print("   ✓ 通過\n")
    
    # 清理
    if os.path.exists(test_file):
        os.remove(test_file)
    
    print("=== 所有測試通過 ✓ ===")


if __name__ == "__main__":
    test_basic_functions()
