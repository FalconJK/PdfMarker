#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
PDF書籤批量添加工具
使用方法:
1. 安裝必要的庫: pip install PyPDF2
2. 準備書籤數據文件 (bookmark_data.txt)，格式如下:
   層級 標題 頁碼
   例如:
   1 第一章 標題 1
   2 1.1 小節 2
   1 第二章 5
3. 運行: python pdf_bookmark_tool.py input.pdf output.pdf bookmark_data.txt [頁碼偏移]
   頁碼偏移: 可選參數，表示目錄頁碼與PDF實際頁碼的差值
   例如，如果目錄中的第1頁實際是PDF的第5頁，則偏移值為4
"""

import sys
import os
import re
from PyPDF2 import PdfReader, PdfWriter

def parse_bookmark_file(bookmark_file, page_offset=0):
    """解析書籤數據文件
    
    Args:
        bookmark_file: 書籤數據文件路徑
        page_offset: 頁碼偏移值，目錄頁碼與PDF實際頁碼的差值
    """
    bookmarks = []
    with open(bookmark_file, 'r', encoding='utf-8') as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith('#'):
                continue
            
            # 改進的解析方式：先獲取第一個數字作為層級
            match = re.match(r'^(\d+)\s+(.+?)\s+(\d+)$', line)
            if match:
                level = int(match.group(1))
                title = match.group(2)
                page = int(match.group(3))
                
                # 應用頁碼偏移，並將頁碼調整為從0開始（PDF內部頁碼從0開始）
                adjusted_page = page + page_offset - 1
                
                bookmarks.append({
                    'level': level,
                    'title': title,
                    'page': adjusted_page
                })
            else:
                print(f"警告: 忽略格式不正確的行: {line}")
                
    return bookmarks

def create_bookmark_tree(bookmarks):
    """將平面書籤列表轉換為樹形結構"""
    root = []
    parents = [root]
    
    prev_level = 0
    for bookmark in bookmarks:
        level = bookmark['level']
        
        # 調整父節點列表以匹配當前級別
        if level > prev_level:
            # 下一級，將上一個書籤作為新的父節點
            if len(parents[-1]) > 0:
                last_item = parents[-1][-1]
                if 'children' not in last_item:
                    last_item['children'] = []
                parents.append(last_item['children'])
        elif level < prev_level:
            # 返回上一級，移除多餘的父節點
            for _ in range(prev_level - level):
                parents.pop()
        
        # 添加當前書籤到適當的父節點
        parents[-1].append({
            'title': bookmark['title'],
            'page': bookmark['page']
        })
        
        prev_level = level
    
    return root

def add_bookmarks_to_pdf(input_pdf, output_pdf, bookmark_file, page_offset=0):
    """向PDF添加書籤
    
    Args:
        input_pdf: 輸入PDF文件路徑
        output_pdf: 輸出PDF文件路徑
        bookmark_file: 書籤數據文件路徑
        page_offset: 頁碼偏移值，目錄頁碼與PDF實際頁碼的差值
    """
    try:
        # 解析書籤數據
        bookmarks = parse_bookmark_file(bookmark_file, page_offset)
        if not bookmarks:
            print("錯誤: 書籤數據為空")
            return False
            
        # 創建書籤樹
        bookmark_tree = create_bookmark_tree(bookmarks)
        
        # 讀取PDF
        reader = PdfReader(input_pdf)
        writer = PdfWriter()
        
        # 檢查頁碼是否超出範圍
        total_pages = len(reader.pages)
        for bookmark in bookmarks:
            if bookmark['page'] < 0 or bookmark['page'] >= total_pages:
                print(f"警告: 書籤 '{bookmark['title']}' 的頁碼 {bookmark['page']+1} 超出PDF範圍 (1-{total_pages})")
        
        # 複製所有頁面
        for page in reader.pages:
            writer.add_page(page)
            
        # 遞歸添加書籤 - 使用新的API
        def add_bookmarks_recursively(parent, bookmarks):
            for bookmark in bookmarks:
                # 確保頁碼在有效範圍內
                page_num = max(0, min(bookmark['page'], total_pages-1))
                
                if 'children' in bookmark:
                    # 創建父書籤
                    new_parent = writer.add_outline_item(
                        bookmark['title'], 
                        page_num,
                        parent=parent
                    )
                    # 遞歸添加子書籤
                    add_bookmarks_recursively(new_parent, bookmark['children'])
                else:
                    # 添加葉子書籤
                    writer.add_outline_item(
                        bookmark['title'], 
                        page_num,
                        parent=parent
                    )
        
        # 添加書籤 - 從根節點開始
        for bookmark in bookmark_tree:
            page_num = max(0, min(bookmark['page'], total_pages-1))
            
            if 'children' in bookmark:
                # 創建頂層書籤
                parent = writer.add_outline_item(bookmark['title'], page_num)
                # 遞歸添加子書籤
                add_bookmarks_recursively(parent, bookmark['children'])
            else:
                # 添加頂層書籤（無子書籤）
                writer.add_outline_item(bookmark['title'], page_num)
        
        # 保存PDF
        with open(output_pdf, 'wb') as f:
            writer.write(f)
            
        print(f"成功: 已將書籤添加到 {output_pdf}")
        if page_offset != 0:
            print(f"已應用頁碼偏移: {page_offset} (目錄頁碼 + {page_offset} = PDF實際頁碼)")
        return True
        
    except Exception as e:
        print(f"錯誤: {str(e)}")
        return False

def main():
    """主函數"""
    if len(sys.argv) < 4 or len(sys.argv) > 5:
        print("使用方法: python pdf_bookmark_tool.py input.pdf output.pdf bookmark_data.txt [頁碼偏移]")
        print("頁碼偏移: 可選參數，表示目錄頁碼與PDF實際頁碼的差值")
        print("例如，如果目錄中的第1頁實際是PDF的第5頁，則偏移值為4")
        return
        
    input_pdf = sys.argv[1]
    output_pdf = sys.argv[2]
    bookmark_file = sys.argv[3]
    
    # 解析頁碼偏移參數（如果提供）
    page_offset = 0
    if len(sys.argv) == 5:
        try:
            page_offset = int(sys.argv[4])
        except ValueError:
            print(f"錯誤: 頁碼偏移必須是整數，而不是 '{sys.argv[4]}'")
            return
    
    # 檢查文件是否存在
    if not os.path.exists(input_pdf):
        print(f"錯誤: 輸入PDF文件不存在: {input_pdf}")
        return
        
    if not os.path.exists(bookmark_file):
        print(f"錯誤: 書籤數據文件不存在: {bookmark_file}")
        return
        
    add_bookmarks_to_pdf(input_pdf, output_pdf, bookmark_file, page_offset)

if __name__ == "__main__":
    main()