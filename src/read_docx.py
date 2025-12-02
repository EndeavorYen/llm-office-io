"""
Read and extract text from Word documents
"""

from typing import List, Optional
import zipfile
import xml.etree.ElementTree as ET
import sys
import os


def read_docx(file_path: str) -> Optional[List[str]]:
    """讀取 Word 文檔並提取所有文字
    
    Args:
        file_path: Word 文檔路徑
        
    Returns:
        Optional[List[str]]: 文字行列表，失敗時返回 None
    """
    if not os.path.exists(file_path):
        print(f"✗ 檔案不存在: {file_path}")
        return None
    
    if not file_path.endswith('.docx'):
        print(f"✗ 不支援的檔案格式: {file_path}")
        return None

    try:
        with zipfile.ZipFile(file_path) as docx:
            xml_content = docx.read('word/document.xml')
            tree = ET.fromstring(xml_content)
            
            # Namespaces in docx xml
            namespaces = {
                'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
            }
            
            text = []
            # Find all paragraphs
            for p in tree.findall('.//w:p', namespaces):
                # Find all runs in paragraph
                p_text = []
                for r in p.findall('.//w:r', namespaces):
                    for t in r.findall('.//w:t', namespaces):
                        if t.text:
                            p_text.append(t.text)
                text.append(''.join(p_text))
            
            return text
            
    except zipfile.BadZipFile:
        print(f"✗ 無效的 ZIP 檔案: {file_path}")
        return None
    except ET.ParseError as e:
        print(f"✗ XML 解析錯誤: {e}")
        return None
    except KeyError:
        print(f"✗ 缺少 word/document.xml: {file_path}")
        return None
    except Exception as e:
        print(f"✗ 讀取檔案時發生錯誤: {e}")
        return None


def main() -> None:
    """主函數"""
    if len(sys.argv) < 2:
        print("用法: python read_docx.py <檔案路徑>")
        sys.exit(1)
    
    text_lines = read_docx(sys.argv[1])
    if text_lines is not None:
        print('\n'.join(text_lines))
    else:
        sys.exit(1)


if __name__ == "__main__":
    main()
