#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
文件夹图片到PDF文档工具使用示例
"""

import os
from pathlib import Path
from folder_to_pdf import create_pdf_document


def main():
    """示例用法"""
    print("=== 文件夹图片到PDF文档工具 ===\n")
    
    # 示例1：使用demo_images文件夹
    demo_folder = Path("demo_images")
    if demo_folder.exists():
        print(f"处理演示文件夹: {demo_folder}")
        
        # 创建PDF文档
        output_file = "演示图片文档.pdf"
        success = create_pdf_document(demo_folder, output_file)
        
        if success:
            print(f"\n✓ 已创建PDF文档: {output_file}")
        else:
            print("\n✗ 创建失败")
    else:
        print("演示文件夹 demo_images 不存在")
    
    # 示例2：手动指定目录
    print("\n" + "="*50)
    print("如果要处理其他目录，请使用以下命令:")
    print("python folder_to_pdf.py <文件夹路径> -o <输出文件名.pdf>")
    print("\n示例:")
    print("python folder_to_pdf.py /path/to/your/folder -o 我的图片文档.pdf")
    
    print("\n支持的图片格式:")
    print("jpg, jpeg, png, gif, bmp, tiff, webp")


if __name__ == "__main__":
    main() 