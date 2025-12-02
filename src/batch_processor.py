#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Batch Processor for Office Documents
批次處理器 - 一次處理多個 Office 文檔
"""

from typing import List, Dict, Optional
import os
import sys
import glob
import shutil
import argparse
from pathlib import Path

try:
    from tqdm import tqdm
    HAS_TQDM = True
except ImportError:
    HAS_TQDM = False
    print("⚠ 建議安裝 tqdm 以顯示進度: pip install tqdm")

from .word_editor import WordEditor
from .ppt_editor import PPTEditor
from .excel_editor import ExcelEditor
from .constants import SUCCESS_SYMBOL, ERROR_SYMBOL, WARNING_SYMBOL


class BatchProcessor:
    """批次處理器類"""
    
    EDITOR_MAP = {
        '.docx': WordEditor,
        '.pptx': PPTEditor,
        '.xlsx': ExcelEditor,
    }
    
    def __init__(self, file_pattern: str, recursive: bool = False):
        """初始化批次處理器
        
        Args:
            file_pattern: 檔案模式 (如 "*.docx", "reports/*.xlsx")
            recursive: 是否遞迴搜尋子目錄
        """
        self.file_pattern = file_pattern
        self.recursive = recursive
        self.files = self._find_files()
        
    def _find_files(self) -> List[str]:
        """尋找符合模式的檔案"""
        if self.recursive:
            # 遞迴搜尋
            pattern = f"**/{self.file_pattern}"
            files = list(glob.glob(pattern, recursive=True))
        else:
            files = list(glob.glob(self.file_pattern))
        
        # 過濾出支援的檔案類型
        supported_files = [
            f for f in files 
            if any(f.endswith(ext) for ext in self.EDITOR_MAP.keys())
        ]
        
        return supported_files
    
    def process_command(
        self, 
        command: str, 
        args: List[str],
        output_dir: Optional[str] = None,
        backup: bool = False
    ) -> Dict[str, bool]:
        """對所有檔案執行相同命令
        
        Args:
            command: 命令名稱 (如 "replace")
            args: 命令參數
            output_dir: 輸出目錄
            backup: 是否備份原檔案
            
        Returns:
            Dict[str, bool]: {檔案路徑: 是否成功}
        """
        if not self.files:
            print(f"{ERROR_SYMBOL} 找不到符合模式的檔案: {self.file_pattern}")
            return {}
        
        print(f"\n找到 {len(self.files)} 個檔案")
        if len(self.files) > 10:
            print(f"前 10 個: {', '.join([os.path.basename(f) for f in self.files[:10]])}")
            print(f"... 還有 {len(self.files) - 10} 個檔案")
        else:
            for f in self.files:
                print(f"  - {f}")
        
        print(f"\n執行命令: {command} {' '.join(args)}\n")
        
        results = {}
        iterator = tqdm(self.files, desc="處理檔案") if HAS_TQDM else self.files
        
        for filepath in iterator:
            if not HAS_TQDM:
                print(f"\n處理: {filepath}")
            
            try:
                success = self._process_single_file(
                    filepath, command, args, output_dir, backup
                )
                results[filepath] = success
            except Exception as e:
                print(f"{ERROR_SYMBOL} {filepath}: {e}")
                results[filepath] = False
        
        # 顯示結果統計
        self._print_summary(results)
        
        return results
    
    def _process_single_file(
        self,
        filepath: str,
        command: str,
        args: List[str],
        output_dir: Optional[str],
        backup: bool
    ) -> bool:
        """處理單個檔案"""
        # 備份原檔案
        if backup:
            backup_path = f"{filepath}.bak"
            shutil.copy2(filepath, backup_path)
        
        # 獲取對應的編輯器
        ext = Path(filepath).suffix
        editor_class = self.EDITOR_MAP.get(ext)
        
        if not editor_class:
            print(f"{WARNING_SYMBOL} 不支援的檔案類型: {ext}")
            return False
        
        # 創建編輯器實例
        editor = editor_class(filepath)
        
        # 執行命令
        success = self._execute_command(editor, command, args)
        
        if success:
            # 儲存檔案
            if output_dir:
                os.makedirs(output_dir, exist_ok=True)
                output_path = os.path.join(output_dir, os.path.basename(filepath))
                
                # 處理檔名衝突
                if os.path.exists(output_path) and output_path != filepath:
                    base, ext = os.path.splitext(output_path)
                    counter = 1
                    while os.path.exists(f"{base}_{counter}{ext}"):
                        counter += 1
                    output_path = f"{base}_{counter}{ext}"
                    print(f"{WARNING_SYMBOL} 檔名衝突，已重新命名為: {os.path.basename(output_path)}")
            else:
                output_path = filepath
            
            editor.save(output_path)
        
        return success
    
    def _execute_command(self, editor, command: str, args: List[str]) -> bool:
        """執行特定命令"""
        try:
            if command == "replace":
                if len(args) < 2:
                    print(f"{ERROR_SYMBOL} replace 需要 2 個參數: old_text new_text")
                    return False
                old_text, new_text = args[0], args[1]
                count = editor.replace_text(old_text, new_text)
                return count > 0
            
            elif command == "delete":
                if len(args) < 1:
                    print(f"{ERROR_SYMBOL} delete 需要 1 個參數: search_text")
                    return False
                if hasattr(editor, 'delete_paragraph'):
                    return editor.delete_paragraph(args[0])
                else:
                    print(f"{WARNING_SYMBOL} 此編輯器不支援 delete 命令")
                    return False
            
            else:
                print(f"{ERROR_SYMBOL} 不支援的命令: {command}")
                return False
                
        except Exception as e:
            print(f"{ERROR_SYMBOL} 執行命令失敗: {e}")
            return False
    
    def _print_summary(self, results: Dict[str, bool]) -> None:
        """顯示處理結果統計"""
        total = len(results)
        success_count = sum(1 for v in results.values() if v)
        fail_count = total - success_count
        
        print(f"\n{'='*50}")
        print(f"處理完成!")
        print(f"  總數: {total}")
        print(f"  {SUCCESS_SYMBOL} 成功: {success_count}")
        if fail_count > 0:
            print(f"  {ERROR_SYMBOL} 失敗: {fail_count}")
        print(f"{'='*50}\n")


def main() -> None:
    """主函數"""
    parser = argparse.ArgumentParser(
        description='批次處理 Office 文檔',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog='''
範例:
  # 批次替換所有 Word 文檔
  python batch_processor.py "*.docx" replace "2024" "2025"
  
  # 批次處理特定目錄的 Excel 檔案
  python batch_processor.py "reports/*.xlsx" replace "舊值" "新值"
  
  # 遞迴處理所有子目錄
  python batch_processor.py "*.pptx" replace "Draft" "Final" --recursive
  
  # 輸出到指定目錄並備份
  python batch_processor.py "*.docx" replace "A" "B" --output out/ --backup
        '''
    )
    
    parser.add_argument('pattern', help='檔案模式 (如 "*.docx", "data/*.xlsx")')
    parser.add_argument('command', help='命令 (replace, delete)')
    parser.add_argument('args', nargs='+', help='命令參數')
    parser.add_argument('--recursive', '-r', action='store_true', help='遞迴搜尋子目錄')
    parser.add_argument('--output', '-o', help='輸出目錄')
    parser.add_argument('--backup', '-b', action='store_true', help='備份原檔案')
    
    args_parsed = parser.parse_args()
    
    # 創建批次處理器
    processor = BatchProcessor(args_parsed.pattern, args_parsed.recursive)
    
    # 執行命令
    results = processor.process_command(
        args_parsed.command,
        args_parsed.args,
        args_parsed.output,
        args_parsed.backup
    )
    
    # 根據結果設定退出碼
    if all(results.values()):
        sys.exit(0)
    elif any(results.values()):
        sys.exit(1)  # 部分成功
    else:
        sys.exit(2)  # 全部失敗


if __name__ == '__main__':
    main()
