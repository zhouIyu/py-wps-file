#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PDF图片提取工具
从PDF文件中提取所有图片并保存到指定目录
"""

import os
import sys
import fitz  # PyMuPDF
import argparse
from pathlib import Path


def extract_images_from_pdf(pdf_path, output_dir="extracted_images"):
    """
    从PDF文件中提取所有图片
    
    Args:
        pdf_path (str): PDF文件路径
        output_dir (str): 输出目录路径
    """
    # 创建输出目录
    output_path = Path(output_dir)
    output_path.mkdir(exist_ok=True)
    
    try:
        # 打开PDF文件
        pdf_document = fitz.open(pdf_path)
        print(f"正在处理PDF文件: {pdf_path}")
        print(f"总页数: {len(pdf_document)}")
        
        image_count = 0
        
        # 遍历每一页
        for page_num in range(len(pdf_document)):
            page = pdf_document[page_num]
            
            # 获取页面中的图片列表
            image_list = page.get_images()
            
            if image_list:
                print(f"页面 {page_num + 1} 找到 {len(image_list)} 张图片")
            
            # 提取每张图片
            for img_index, img in enumerate(image_list):
                try:
                    # 获取图片的xref
                    xref = img[0]
                    
                    # 提取图片数据
                    base_image = pdf_document.extract_image(xref)
                    image_bytes = base_image["image"]
                    image_ext = base_image["ext"]
                    
                    # 生成文件名
                    image_filename = f"page_{page_num + 1}_img_{img_index + 1}.{image_ext}"
                    image_path = output_path / image_filename
                    
                    # 保存图片
                    with open(image_path, "wb") as image_file:
                        image_file.write(image_bytes)
                    
                    image_count += 1
                    print(f"  ✓ 保存图片: {image_filename}")
                    
                except Exception as e:
                    print(f"  ✗ 提取图片失败 (页面 {page_num + 1}, 图片 {img_index + 1}): {e}")
        
        pdf_document.close()
        
        print(f"\n提取完成! 总共提取了 {image_count} 张图片")
        print(f"图片保存在: {output_path.absolute()}")
        
        return image_count
        
    except Exception as e:
        print(f"处理PDF文件时出错: {e}")
        return 0


def main():
    """主函数"""
    parser = argparse.ArgumentParser(description="从PDF文件中提取所有图片")
    parser.add_argument("pdf_file", help="PDF文件路径")
    parser.add_argument("-o", "--output", default="extracted_images", 
                       help="输出目录 (默认: extracted_images)")
    
    args = parser.parse_args()
    
    # 检查PDF文件是否存在
    if not os.path.exists(args.pdf_file):
        print(f"错误: PDF文件不存在: {args.pdf_file}")
        sys.exit(1)
    
    # 检查文件是否为PDF
    if not args.pdf_file.lower().endswith('.pdf'):
        print(f"警告: 文件可能不是PDF格式: {args.pdf_file}")
    
    # 提取图片
    extracted_count = extract_images_from_pdf(args.pdf_file, args.output)
    
    if extracted_count == 0:
        print("没有找到图片或提取失败")
        sys.exit(1)


if __name__ == "__main__":
    main() 