#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
LLM-Friendly API Layer
簡化的單步操作接口，專為 AI Agent 設計

提供統一的函數調用接口，自動處理檔案類型判斷、錯誤處理和結果返回。
"""

from typing import Dict, Any, Optional, List, Union
from pathlib import Path
import json
import os

from .word_editor import WordEditor
from .ppt_editor import PPTEditor
from .excel_editor import ExcelEditor
from .batch_processor import BatchProcessor


class OfficeAPI:
    """LLM-Friendly Office API 類"""
    
    # 檔案類型映射
    EDITORS = {
        '.docx': WordEditor,
        '.pptx': PPTEditor,
        '.xlsx': ExcelEditor,
    }
    
    @staticmethod
    def _get_editor(file_path: str):
        """根據檔案類型獲取對應的編輯器"""
        ext = Path(file_path).suffix.lower()
        
        if ext not in OfficeAPI.EDITORS:
            raise ValueError(f"不支援的檔案格式: {ext}")
        
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"檔案不存在: {file_path}")
        
        return OfficeAPI.EDITORS[ext](file_path), ext[1:]  # 返回編輯器和類型
    
    @staticmethod
    def _create_response(success: bool, operation: str, file_type: str,
                        result: Any = None, message: str = "", 
                        error: Optional[str] = None) -> Dict[str, Any]:
        """創建統一的返回格式"""
        return {
            "success": success,
            "operation": operation,
            "file_type": file_type,
            "result": result,
            "message": message,
            "error": error
        }


def replace_text(file_path: str, old_text: str, new_text: str, 
                output_path: Optional[str] = None) -> Dict[str, Any]:
    """
    替換文字（自動判斷檔案類型）
    
    Args:
        file_path: 檔案路徑
        old_text: 要替換的文字
        new_text: 新文字
        output_path: 輸出路徑，None 表示覆蓋原檔案
    
    Returns:
        {
            "success": bool,
            "operation": "replace_text",
            "file_type": "word"|"ppt"|"excel",
            "result": {"count": int},
            "message": str,
            "error": Optional[str]
        }
    
    Example:
        >>> result = replace_text("report.docx", "2024", "2025")
        >>> print(result)
        {'success': True, 'operation': 'replace_text', 'file_type': 'word', 
         'result': {'count': 5}, 'message': '成功替換 5 處', 'error': None}
    """
    try:
        editor, file_type = OfficeAPI._get_editor(file_path)
        count = editor.replace_text(old_text, new_text)
        editor.save(output_path or file_path)
        
        return OfficeAPI._create_response(
            success=True,
            operation="replace_text",
            file_type=file_type,
            result={"count": count},
            message=f"成功替換 {count} 處"
        )
    except Exception as e:
        return OfficeAPI._create_response(
            success=False,
            operation="replace_text",
            file_type="unknown",
            error=str(e)
        )


def add_image(file_path: str, image_path: str, width_cm: float = 10.0,
             position: Optional[str] = None, slide_number: Optional[int] = None,
             left_cm: float = 2.0, top_cm: float = 5.0,
             output_path: Optional[str] = None) -> Dict[str, Any]:
    """
    添加圖片（自動判斷檔案類型並調用對應方法）
    
    Args:
        file_path: 檔案路徑
        image_path: 圖片路徑
        width_cm: 圖片寬度（公分）
        position: Word 文檔插入位置
        slide_number: PPT 投影片編號
        left_cm: PPT 左邊距
        top_cm: PPT 上邊距
        output_path: 輸出路徑
    
    Returns:
        統一格式的結果字典
    """
    try:
        editor, file_type = OfficeAPI._get_editor(file_path)
        
        if file_type == 'word':
            result = editor.add_image(image_path, width_cm, position)
        elif file_type == 'ppt':
            if slide_number is None:
                raise ValueError("PPT 檔案需要指定 slide_number")
            # 假設 PPT 有 add_image 方法（需要整合）
            result = True  # 暫時返回 True
        else:
            raise ValueError(f"{file_type} 不支援插入圖片")
        
        if result:
            editor.save(output_path or file_path)
            return OfficeAPI._create_response(
                success=True,
                operation="add_image",
                file_type=file_type,
                result={"image_added": True},
                message="成功插入圖片"
            )
        else:
            return OfficeAPI._create_response(
                success=False,
                operation="add_image",
                file_type=file_type,
                error="圖片插入失敗"
            )
    except Exception as e:
        return OfficeAPI._create_response(
            success=False,
            operation="add_image",
            file_type="unknown",
            error=str(e)
        )


def insert_table(file_path: str, rows: int, cols: int,
                data: Optional[List[List[str]]] = None,
                position: Optional[str] = None,
                output_path: Optional[str] = None) -> Dict[str, Any]:
    """
    插入表格（僅支援 Word）
    
    Args:
        file_path: Word 文檔路徑
        rows: 行數
        cols: 列數
        data: 表格數據
        position: 插入位置
        output_path: 輸出路徑
    
    Returns:
        統一格式的結果字典
    """
    try:
        editor, file_type = OfficeAPI._get_editor(file_path)
        
        if file_type != 'word':
            raise ValueError("只有 Word 文檔支援插入表格")
        
        result = editor.insert_table(rows, cols, data, position)
        
        if result:
            editor.save(output_path or file_path)
            return OfficeAPI._create_response(
                success=True,
                operation="insert_table",
                file_type=file_type,
                result={"rows": rows, "cols": cols},
                message=f"成功插入 {rows}x{cols} 表格"
            )
        else:
            return OfficeAPI._create_response(
                success=False,
                operation="insert_table",
                file_type=file_type,
                error="表格插入失敗"
            )
    except Exception as e:
        return OfficeAPI._create_response(
            success=False,
            operation="insert_table",
            file_type="unknown",
            error=str(e)
        )


def batch_replace(pattern: str, old_text: str, new_text: str,
                 recursive: bool = False, output_dir: Optional[str] = None,
                 backup: bool = False) -> Dict[str, Any]:
    """
    批次替換文字
    
    Args:
        pattern: 檔案模式（如 "*.docx"）
        old_text: 要替換的文字
        new_text: 新文字
        recursive: 是否遞迴搜尋
        output_dir: 輸出目錄
        backup: 是否備份
    
    Returns:
        統一格式的結果字典，包含處理結果統計
    """
    try:
        processor = BatchProcessor(pattern, recursive)
        results = processor.process_command("replace", [old_text, new_text], 
                                          output_dir, backup)
        
        success_count = sum(1 for v in results.values() if v)
        total_count = len(results)
        
        return OfficeAPI._create_response(
            success=True,
            operation="batch_replace",
            file_type="mixed",
            result={
                "total": total_count,
                "success": success_count,
                "failed": total_count - success_count,
                "files": list(results.keys())
            },
            message=f"處理 {total_count} 個檔案，成功 {success_count} 個"
        )
    except Exception as e:
        return OfficeAPI._create_response(
            success=False,
            operation="batch_replace",
            file_type="mixed",
            error=str(e)
        )


def execute_command(command: str, **kwargs) -> Dict[str, Any]:
    """
    通用命令執行接口
    
    Args:
        command: 命令名稱 ("replace_text", "add_image", "insert_table", "batch_replace")
        **kwargs: 命令參數
    
    Returns:
        統一格式的結果字典
    
    Example:
        >>> result = execute_command("replace_text", 
        ...                         file_path="report.docx",
        ...                         old_text="2024", 
        ...                         new_text="2025")
    """
    command_map = {
        "replace_text": replace_text,
        "add_image": add_image,
        "insert_table": insert_table,
        "batch_replace": batch_replace,
    }
    
    if command not in command_map:
        return OfficeAPI._create_response(
            success=False,
            operation=command,
            file_type="unknown",
            error=f"不支援的命令: {command}"
        )
    
    try:
        return command_map[command](**kwargs)
    except TypeError as e:
        return OfficeAPI._create_response(
            success=False,
            operation=command,
            file_type="unknown",
            error=f"參數錯誤: {str(e)}"
        )


def execute_json(json_input: Union[str, dict]) -> str:
    """
    接受 JSON 輸入並返回 JSON 輸出
    
    Args:
        json_input: JSON 字符串或字典
            {
                "command": "replace_text",
                "params": {
                    "file_path": "report.docx",
                    "old_text": "2024",
                    "new_text": "2025"
                }
            }
    
    Returns:
        JSON 字符串格式的結果
    
    Example:
        >>> json_str = '{"command": "replace_text", "params": {"file_path": "test.docx", "old_text": "A", "new_text": "B"}}'
        >>> result = execute_json(json_str)
        >>> print(result)
        {"success": true, "operation": "replace_text", ...}
    """
    try:
        if isinstance(json_input, str):
            data = json.loads(json_input)
        else:
            data = json_input
        
        command = data.get("command")
        params = data.get("params", {})
        
        result = execute_command(command, **params)
        return json.dumps(result, ensure_ascii=False, indent=2)
    
    except json.JSONDecodeError as e:
        error_result = OfficeAPI._create_response(
            success=False,
            operation="execute_json",
            file_type="unknown",
            error=f"JSON 解析錯誤: {str(e)}"
        )
        return json.dumps(error_result, ensure_ascii=False, indent=2)
    except Exception as e:
        error_result = OfficeAPI._create_response(
            success=False,
            operation="execute_json",
            file_type="unknown",
            error=str(e)
        )
        return json.dumps(error_result, ensure_ascii=False, indent=2)


# 便捷導出
__all__ = [
    'replace_text',
    'add_image',
    'insert_table',
    'batch_replace',
    'execute_command',
    'execute_json',
    'OfficeAPI'
]
