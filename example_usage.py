#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PDF图片提取示例
演示如何在代码中使用extract_pdf_images函数
"""

from extract_pdf_images import extract_images_from_pdf
import os


def demo_extract():
    """演示函数用法"""
    print("=== PDF图片提取工具演示 ===\n")
    
    # 示例PDF文件路径（请替换为实际的PDF文件路径）
    pdf_file = "sample.pdf"
    output_dir = "my_images"
    
    # 检查文件是否存在
    if not os.path.exists(pdf_file):
        print(f"请将您的PDF文件命名为 '{pdf_file}' 并放在当前目录中")
        print("或者修改上面的 pdf_file 变量为您的PDF文件路径")
        return
    
    print(f"准备从 {pdf_file} 中提取图片...")
    print(f"输出目录: {output_dir}\n")
    
    # 调用提取函数
    count = extract_images_from_pdf(pdf_file, output_dir)
    
    if count > 0:
        print(f"\n✅ 成功提取 {count} 张图片!")
        print(f"📁 图片保存在: {os.path.abspath(output_dir)}")
    else:
        print("\n❌ 没有找到图片或提取失败")


if __name__ == "__main__":
    demo_extract() 