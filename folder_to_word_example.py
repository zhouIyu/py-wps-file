#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
文件夹图片到Word文档工具使用示例
"""

import os
from pathlib import Path
from folder_to_word import create_word_document


def main():
    """示例用法"""
    print("=== 文件夹图片到Word文档工具 ===\n")
    
    # 示例1：使用当前目录作为源文件夹
    current_dir = Path.cwd()
    print(f"当前目录: {current_dir}")
    
    # 查看当前目录的子目录
    subdirs = [item for item in current_dir.iterdir() if item.is_dir()]
    if subdirs:
        print(f"发现子目录: {[d.name for d in subdirs]}")
        
        # 创建Word文档
        output_file = "当前目录图片文档.docx"
        success = create_word_document(current_dir, output_file)
        
        if success:
            print(f"\n✓ 已创建Word文档: {output_file}")
        else:
            print("\n✗ 创建失败")
    else:
        print("当前目录下没有子目录")
    
    # 示例2：手动指定目录
    print("\n" + "="*50)
    print("如果要处理其他目录，请使用以下命令:")
    print("python folder_to_word.py <文件夹路径> -o <输出文件名.docx>")
    print("\n示例:")
    print("python folder_to_word.py /path/to/your/folder -o 我的图片文档.docx")
    
    print("\n支持的图片格式:")
    print("jpg, jpeg, png, gif, bmp, tiff, webp")


if __name__ == "__main__":
    main() 