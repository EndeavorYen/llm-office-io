#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Interactive Word Document Editor
直接操作 Word 文檔的互動式編輯工具
"""

from typing import Optional, List
import os
import sys
import argparse

from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

from .constants import (
    MAX_PREVIEW_LENGTH,
    MAX_TEXT_DISPLAY,
    SUCCESS_SYMBOL,
    ERROR_SYMBOL,
    WORD_EXTENSION,
    DEFAULT_HEADING_LEVEL
)


class WordEditor:
    """Word 文檔編輯器類"""
    
    def __init__(self, filepath: str) -> None:
        """初始化 Word 編輯器
        
        Args:
            filepath: Word 文檔路徑
            
        Raises:
            FileNotFoundError: 當檔案不存在時
            ValueError: 當檔案格式不支援時
            RuntimeError: 當無法開啟文檔時
        """
        if not os.path.exists(filepath):
            raise FileNotFoundError(f"檔案不存在: {filepath}")
        
        if not filepath.endswith(WORD_EXTENSION):
            raise ValueError(f"不支援的檔案格式，需要 {WORD_EXTENSION}: {filepath}")
        
        try:
            self.filepath = filepath
            self.doc = Document(filepath)
        except Exception as e:
            raise RuntimeError(f"無法開啟文檔: {e}") from e
    
    def save(self, output_path: Optional[str] = None) -> None:
        """儲存文檔
        
        Args:
            output_path: 輸出路徑，None 表示覆蓋原檔案
        """
        save_path = output_path or self.filepath
        try:
            self.doc.save(save_path)
            print(f"{SUCCESS_SYMBOL} 文檔已儲存: {save_path}")
        except Exception as e:
            print(f"{ERROR_SYMBOL} 儲存失敗: {e}")
            raise
    
    def list_structure(self) -> None:
        """列出文檔結構（標題和段落）"""
        print("\n=== 文檔結構 ===\n")
        for i, para in enumerate(self.doc.paragraphs):
            if para.style.name.startswith('Heading'):
                level = para.style.name.replace('Heading ', '')
                indent = "  " * (int(level) - 1) if level.isdigit() else ""
                print(f"[{i}] {indent}📌 {para.text[:MAX_PREVIEW_LENGTH]}")
            elif para.text.strip():
                preview = para.text[:MAX_PREVIEW_LENGTH].replace('\n', ' ')
                print(f"[{i}]    {preview}")
        print()
    
    def add_paragraph_after(
        self, 
        search_text: str, 
        new_content: str, 
        heading_level: Optional[int] = None
    ) -> bool:
        """在包含特定文字的段落後添加新段落
        
        Args:
            search_text: 搜尋文字
            new_content: 新內容
            heading_level: 標題層級 (1-9)，None 表示普通段落
            
        Returns:
            bool: 是否找到並添加成功
        """
        if not search_text:
            print(f"{ERROR_SYMBOL} 搜尋文字不能為空")
            return False
            
        found = False
        for i, para in enumerate(self.doc.paragraphs):
            if search_text in para.text:
                # 在找到的段落後插入
                p = para._element
                parent = p.getparent()
                
                # 創建新段落
                if heading_level:
                    new_para = self.doc.add_heading(new_content, level=heading_level)
                else:
                    new_para = self.doc.add_paragraph(new_content)
                
                # 移動到正確位置
                parent.insert(parent.index(p) + 1, new_para._element)
                
                preview = para.text[:50]
                print(f"{SUCCESS_SYMBOL} 已在「{preview}...」後添加內容")
                found = True
                break
        
        if not found:
            print(f"{ERROR_SYMBOL} 找不到包含「{search_text}」的段落")
            
        return found
    
    def replace_text(self, old_text: str, new_text: str, count: int = -1) -> int:
        """替換文字（支援段落和表格）
        
        Args:
            old_text: 要替換的文字
            new_text: 新文字
            count: 替換次數，-1 表示全部替換
            
        Returns:
            int: 實際替換的次數
        """
        if not old_text:
            print(f"{ERROR_SYMBOL} 要替換的文字不能為空")
            return 0
            
        replaced_count = 0
        
        # 替換段落中的文字
        for para in self.doc.paragraphs:
            if old_text in para.text:
                for run in para.runs:
                    if old_text in run.text:
                        run.text = run.text.replace(old_text, new_text, 1 if count > 0 else -1)
                        replaced_count += 1
                        if count > 0 and replaced_count >= count:
                            break
        
        # 替換表格中的文字
        if count < 0 or replaced_count < count:
            for table in self.doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        if old_text in cell.text:
                            for para in cell.paragraphs:
                                for run in para.runs:
                                    if old_text in run.text:
                                        run.text = run.text.replace(
                                            old_text, new_text, 1 if count > 0 else -1
                                        )
                                        replaced_count += 1
                                        if count > 0 and replaced_count >= count:
                                            break
        
        if replaced_count > 0:
            print(f"{SUCCESS_SYMBOL} 已替換 {replaced_count} 處「{old_text}」→「{new_text}」")
        else:
            print(f"{ERROR_SYMBOL} 找不到「{old_text}」")
            
        return replaced_count
    
    def delete_paragraph(self, search_text: str) -> bool:
        """刪除包含特定文字的段落
        
        Args:
            search_text: 搜尋文字
            
        Returns:
            bool: 是否找到並刪除成功
        """
        if not search_text:
            print(f"{ERROR_SYMBOL} 搜尋文字不能為空")
            return False
            
        deleted = False
        for para in self.doc.paragraphs:
            if search_text in para.text:
                p = para._element
                p.getparent().remove(p)
                print(f"{SUCCESS_SYMBOL} 已刪除段落: {para.text[:50]}")
                deleted = True
                break
        
        if not deleted:
            print(f"{ERROR_SYMBOL} 找不到包含「{search_text}」的段落")
            
        return deleted
    
    def insert_after_heading(
        self, 
        heading_text: str, 
        content: str, 
        is_heading: bool = False, 
        heading_level: int = DEFAULT_HEADING_LEVEL
    ) -> bool:
        """在特定標題後插入內容
        
        Args:
            heading_text: 標題文字
            content: 要插入的內容
            is_heading: 插入的內容是否為標題
            heading_level: 標題層級 (1-9)
            
        Returns:
            bool: 是否找到並插入成功
        """
        if not heading_text:
            print(f"{ERROR_SYMBOL} 標題文字不能為空")
            return False
            
        found = False
        for i, para in enumerate(self.doc.paragraphs):
            if para.style.name.startswith('Heading') and heading_text in para.text:
                # 找到標題，在它後面插入
                p = para._element
                parent = p.getparent()
                
                if is_heading:
                    new_para = self.doc.add_heading(content, level=heading_level)
                else:
                    new_para = self.doc.add_paragraph(content)
                
                parent.insert(parent.index(p) + 1, new_para._element)
                
                print(f"{SUCCESS_SYMBOL} 已在標題「{para.text}」後插入內容")
                found = True
                break
        
        if not found:
            print(f"{ERROR_SYMBOL} 找不到標題「{heading_text}」")
            
        return found
    
    def add_bullet_points(self, heading_text: str, bullet_points: List[str]) -> bool:
        """在特定標題後添加多個項目符號
        
        Args:
            heading_text: 標題文字
            bullet_points: 項目列表
            
        Returns:
            bool: 是否找到並添加成功
        """
        if not heading_text:
            print(f"{ERROR_SYMBOL} 標題文字不能為空")
            return False
            
        if not bullet_points:
            print(f"{ERROR_SYMBOL} 項目列表不能為空")
            return False
            
        found = False
        for i, para in enumerate(self.doc.paragraphs):
            if para.style.name.startswith('Heading') and heading_text in para.text:
                p = para._element
                parent = p.getparent()
                insert_pos = parent.index(p) + 1
                
                for bullet in bullet_points:
                    new_para = self.doc.add_paragraph(f"• {bullet}")
                    parent.insert(insert_pos, new_para._element)
                    insert_pos += 1
                
                print(f"{SUCCESS_SYMBOL} 已在「{para.text}」後添加 {len(bullet_points)} 個項目")
                found = True
                break
        
        if not found:
            print(f"{ERROR_SYMBOL} 找不到標題「{heading_text}」")
            
        return found


def main() -> None:
    """主函數"""
    parser = argparse.ArgumentParser(description='Word 文檔互動式編輯器')
    parser.add_argument('file', help='Word 文檔路徑')
    parser.add_argument('--output', '-o', help='輸出文件路徑（不指定則覆蓋原文件）')
    
    subparsers = parser.add_subparsers(dest='command', help='編輯命令')
    
    # list: 列出文檔結構
    subparsers.add_parser('list', help='列出文檔結構')
    
    # replace: 替換文字
    replace_parser = subparsers.add_parser('replace', help='替換文字')
    replace_parser.add_argument('old', help='要替換的文字')
    replace_parser.add_argument('new', help='新文字')
    replace_parser.add_argument('--count', type=int, default=-1, help='替換次數（-1表示全部）')
    
    # add-after: 在段落後添加內容
    add_parser = subparsers.add_parser('add-after', help='在特定段落後添加內容')
    add_parser.add_argument('search', help='搜尋文字')
    add_parser.add_argument('content', help='要添加的內容')
    add_parser.add_argument('--heading', type=int, help='作為標題（指定層級1-3）')
    
    # insert-after-heading: 在標題後插入
    insert_parser = subparsers.add_parser('insert-after-heading', help='在標題後插入內容')
    insert_parser.add_argument('heading', help='標題文字')
    insert_parser.add_argument('content', help='要插入的內容')
    insert_parser.add_argument('--heading-level', type=int, default=DEFAULT_HEADING_LEVEL, 
                              help='作為標題層級')
    insert_parser.add_argument('--is-heading', action='store_true', help='插入的內容是標題')
    
    # delete: 刪除段落
    delete_parser = subparsers.add_parser('delete', help='刪除段落')
    delete_parser.add_argument('search', help='要刪除的段落（搜尋文字）')
    
    # add-bullets: 添加項目符號
    bullets_parser = subparsers.add_parser('add-bullets', help='在標題後添加項目符號')
    bullets_parser.add_argument('heading', help='標題文字')
    bullets_parser.add_argument('bullets', nargs='+', help='項目內容（可多個）')
    
    args = parser.parse_args()
    
    if not args.command:
        parser.print_help()
        return
    
    # 載入文檔
    try:
        editor = WordEditor(args.file)
    except (FileNotFoundError, ValueError, RuntimeError) as e:
        print(f"{ERROR_SYMBOL} {e}")
        sys.exit(1)
    
    # 執行命令
    try:
        if args.command == 'list':
            editor.list_structure()
            return
        
        elif args.command == 'replace':
            editor.replace_text(args.old, args.new, args.count)
        
        elif args.command == 'add-after':
            editor.add_paragraph_after(args.search, args.content, args.heading)
        
        elif args.command == 'insert-after-heading':
            editor.insert_after_heading(args.heading, args.content, 
                                       args.is_heading, args.heading_level)
        
        elif args.command == 'delete':
            editor.delete_paragraph(args.search)
        
        elif args.command == 'add-bullets':
            editor.add_bullet_points(args.heading, args.bullets)
        
        # 儲存
        editor.save(args.output)
        
    except Exception as e:
        print(f"{ERROR_SYMBOL} 操作失敗: {e}")
        sys.exit(1)


if __name__ == '__main__':
    main()
