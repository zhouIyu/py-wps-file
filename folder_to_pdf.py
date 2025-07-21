#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
文件夹图片到PDF文档工具
将文件夹中的二级目录作为标题，并将目录内的图片插入到PDF文档中
"""

import os
import sys
import argparse
from pathlib import Path
import fitz  # PyMuPDF
from PIL import Image
import logging

# 支持的图片格式
SUPPORTED_IMAGE_FORMATS = {'.jpg', '.jpeg', '.png', '.gif', '.bmp', '.tiff', '.webp'}

# 设置日志
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)


def is_image_file(file_path):
    """检查文件是否为支持的图片格式"""
    return file_path.suffix.lower() in SUPPORTED_IMAGE_FORMATS


def get_image_size_for_pdf(img_path, max_width=500, max_height=700):
    """
    获取适合PDF的图片尺寸
    
    Args:
        img_path (Path): 图片路径
        max_width (int): 最大宽度（像素）
        max_height (int): 最大高度（像素）
    
    Returns:
        tuple: (width, height) 调整后的尺寸
    """
    try:
        with Image.open(img_path) as img:
            original_width, original_height = img.size
            
            # 计算缩放比例
            width_ratio = max_width / original_width
            height_ratio = max_height / original_height
            ratio = min(width_ratio, height_ratio, 1.0)  # 不放大，只缩小
            
            new_width = int(original_width * ratio)
            new_height = int(original_height * ratio)
            
            return new_width, new_height
    except Exception as e:
        logger.warning(f"无法获取图片尺寸 {img_path}: {e}")
        return 400, 300  # 默认尺寸


def create_pdf_document(folder_path, output_path="图片文档.pdf"):
    """
    创建PDF文档，将文件夹中的二级目录作为标题，图片插入到对应标题下
    
    Args:
        folder_path (str): 源文件夹路径
        output_path (str): 输出PDF文档路径
    """
    folder_path = Path(folder_path)
    
    if not folder_path.exists():
        logger.error(f"文件夹不存在: {folder_path}")
        return False
    
    if not folder_path.is_dir():
        logger.error(f"路径不是文件夹: {folder_path}")
        return False
    
    # 创建PDF文档
    doc = fitz.open()  # 创建新的PDF文档
    
    # 统计信息
    total_directories = 0
    total_images = 0
    
    # 获取所有二级目录
    subdirectories = []
    try:
        for item in folder_path.iterdir():
            if item.is_dir():
                subdirectories.append(item)
    except PermissionError:
        logger.error(f"没有权限访问文件夹: {folder_path}")
        return False
    
    # 按目录名排序
    subdirectories.sort(key=lambda x: x.name)
    
    if not subdirectories:
        logger.warning(f"在 {folder_path} 中没有找到子目录")
        # 创建一个页面显示没有内容
        page = doc.new_page()
        text = "没有找到子目录"
        page.insert_text((100, 100), text, fontsize=20, color=(0, 0, 0))
        doc.save(output_path)
        doc.close()
        return True
    
    logger.info(f"找到 {len(subdirectories)} 个子目录")
    
    # 先扫描所有目录，收集有图片的目录信息
    valid_directories = []
    for subdir in subdirectories:
        # 获取目录中的所有图片文件
        image_files = []
        try:
            for file_path in subdir.iterdir():
                if file_path.is_file() and is_image_file(file_path):
                    image_files.append(file_path)
        except PermissionError:
            logger.warning(f"没有权限访问目录: {subdir}")
            continue
        
        if image_files:  # 只记录有图片的目录
            valid_directories.append((subdir, image_files))
    
    if not valid_directories:
        logger.warning(f"没有找到包含图片的目录")
        page = doc.new_page()
        text = "没有找到包含图片的目录"
        page.insert_text((100, 100), text, fontsize=20, color=(0, 0, 0))
        doc.save(output_path)
        doc.close()
        return True
    
    # 添加封面页
    page = doc.new_page()
    
    # 添加标题
    title_text = f'{folder_path.name} - 图片文档'
    page.insert_text((200, 100), title_text, fontsize=24, color=(0, 0, 0))
    
    # 添加目录
    page.insert_text((100, 200), '目录', fontsize=20, color=(0, 0, 0))
    
    # 添加目录项
    y_pos = 250
    for i, (subdir, img_files) in enumerate(valid_directories, 1):
        toc_item = f"{i}. {subdir.name} ({len(img_files)}张图片)"
        page.insert_text((120, y_pos), toc_item, fontsize=14, color=(0, 0, 0))
        y_pos += 25
    
    # 遍历每个有效目录
    for subdir, image_files in valid_directories:
        logger.info(f"处理目录: {subdir.name}")
        
        # 按文件名排序
        image_files.sort(key=lambda x: x.name)
        
        total_directories += 1
        
        # 为每个目录创建新页面
        page = doc.new_page()
        
        # 添加目录标题
        page.insert_text((100, 80), subdir.name, fontsize=22, color=(0, 0, 0))
        
        logger.info(f"  找到 {len(image_files)} 张图片")
        
        # 当前页面的y位置
        current_y = 120
        page_height = 800  # 页面高度限制
        
        # 添加图片
        for i, img_path in enumerate(image_files, 1):
            try:
                logger.info(f"    添加图片: {img_path.name}")
                
                # 获取适当的图片尺寸
                img_width, img_height = get_image_size_for_pdf(img_path)
                
                # 检查是否需要新页面
                if current_y + img_height + 100 > page_height:  # 留出底部边距
                    page = doc.new_page()
                    current_y = 80
                
                # 计算居中位置
                page_width = 595  # A4页面宽度（点）
                x_pos = (page_width - img_width) / 2
                
                try:
                    # 插入图片
                    img_rect = fitz.Rect(x_pos, current_y, x_pos + img_width, current_y + img_height)
                    page.insert_image(img_rect, filename=str(img_path))
                    
                    # 在图片下方添加文件名
                    text_y = current_y + img_height + 10
                    page.insert_text((x_pos, text_y), f"图片 {i}: {img_path.name}", 
                                   fontsize=10, color=(0.5, 0.5, 0.5))
                    
                    # 更新位置
                    current_y = text_y + 30
                    total_images += 1
                    
                except Exception as e:
                    logger.error(f"    插入图片失败 {img_path.name}: {e}")
                    # 添加错误说明
                    error_text = f"无法加载图片: {img_path.name}"
                    page.insert_text((100, current_y), error_text, fontsize=12, color=(1, 0, 0))
                    current_y += 30
                
            except Exception as e:
                logger.error(f"    处理图片失败 {img_path.name}: {e}")
                continue
    
    # 保存PDF文档
    try:
        doc.save(output_path)
        doc.close()
        logger.info(f"PDF文档已保存: {Path(output_path).absolute()}")
        logger.info(f"统计信息 - 目录: {total_directories}，图片: {total_images}")
        return True
    except Exception as e:
        logger.error(f"保存PDF文档失败: {e}")
        if doc:
            doc.close()
        return False


def main():
    """主函数"""
    parser = argparse.ArgumentParser(description="将文件夹中的二级目录作为标题，图片插入到PDF文档中")
    parser.add_argument("folder_path", help="源文件夹路径")
    parser.add_argument("-o", "--output", default="图片文档.pdf", 
                       help="输出PDF文档路径 (默认: 图片文档.pdf)")
    
    args = parser.parse_args()
    
    # 检查文件夹是否存在
    if not os.path.exists(args.folder_path):
        print(f"错误: 文件夹不存在: {args.folder_path}")
        sys.exit(1)
    
    # 创建PDF文档
    success = create_pdf_document(args.folder_path, args.output)
    
    if success:
        print(f"\n✓ 成功创建PDF文档: {args.output}")
    else:
        print("✗ 创建PDF文档失败")
        sys.exit(1)


if __name__ == "__main__":
    main() 