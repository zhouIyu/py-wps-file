#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
构建exe文件的脚本
使用PyInstaller将Python脚本打包成可执行文件
"""

import os
import sys
import subprocess
import shutil
from pathlib import Path


def install_dependencies():
    """安装所需依赖"""
    print("正在安装依赖包...")
    try:
        subprocess.run([sys.executable, "-m", "pip", "install", "-r", "requirements.txt"], 
                      check=True, capture_output=True, text=True)
        print("✓ 依赖包安装完成")
        return True
    except subprocess.CalledProcessError as e:
        print(f"✗ 依赖包安装失败: {e}")
        print(f"错误输出: {e.stderr}")
        return False


def build_exe():
    """构建exe文件"""
    print("正在构建exe文件...")
    
    # 确保输出目录存在
    dist_dir = Path("dist")
    build_dir = Path("build")
    
    # 清理之前的构建文件
    if dist_dir.exists():
        shutil.rmtree(dist_dir)
    if build_dir.exists():
        shutil.rmtree(build_dir)
    
    # PyInstaller命令
    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--onefile",  # 生成单个exe文件
        "--windowed",  # 不显示控制台窗口（GUI程序）
        "--name", "文件夹图片转Excel工具",  # exe文件名
        "--icon", "icon.ico" if Path("icon.ico").exists() else None,  # 图标文件（如果存在）
        "--add-data", "README.md;.",  # 添加README文件
        "folder_to_excel.py"  # 主程序文件
    ]
    
    # 移除None项
    cmd = [item for item in cmd if item is not None]
    
    try:
        result = subprocess.run(cmd, check=True, capture_output=True, text=True)
        print("✓ exe文件构建完成")
        
        # 检查生成的文件
        exe_file = dist_dir / "文件夹图片转Excel工具.exe"
        if exe_file.exists():
            file_size = exe_file.stat().st_size / (1024 * 1024)  # MB
            print(f"✓ 生成的exe文件: {exe_file}")
            print(f"  文件大小: {file_size:.2f} MB")
        else:
            print("✗ 未找到生成的exe文件")
            return False
            
        return True
        
    except subprocess.CalledProcessError as e:
        print(f"✗ exe构建失败: {e}")
        print(f"错误输出: {e.stderr}")
        return False


def create_icon():
    """创建简单的图标文件（如果不存在）"""
    icon_path = Path("icon.ico")
    if not icon_path.exists():
        print("创建默认图标...")
        try:
            from PIL import Image, ImageDraw
            
            # 创建32x32的简单图标
            img = Image.new('RGB', (32, 32), color='lightblue')
            draw = ImageDraw.Draw(img)
            
            # 绘制简单的文件夹图标
            draw.rectangle([4, 8, 28, 24], outline='darkblue', width=2)
            draw.rectangle([4, 6, 16, 10], fill='darkblue')
            draw.text((8, 14), 'XLS', fill='darkblue')
            
            img.save(icon_path, format='ICO')
            print(f"✓ 图标文件已创建: {icon_path}")
            return True
            
        except Exception as e:
            print(f"✗ 创建图标失败: {e}")
            return False
    else:
        print(f"✓ 图标文件已存在: {icon_path}")
        return True


def create_build_info():
    """创建构建信息文件"""
    info_file = Path("build_info.txt")
    
    with open(info_file, 'w', encoding='utf-8') as f:
        f.write("文件夹图片转Excel工具 - 构建信息\n")
        f.write("="*50 + "\n\n")
        f.write("功能说明:\n")
        f.write("- 扫描指定文件夹中的二级目录\n")
        f.write("- 将每个目录的名称作为Excel工作表\n")
        f.write("- 将目录中的图片插入到对应的工作表中\n")
        f.write("- 支持的图片格式: JPG, JPEG, PNG, GIF, BMP, TIFF, WEBP\n\n")
        f.write("使用方法:\n")
        f.write("1. 双击运行exe文件\n")
        f.write("2. 选择包含图片的文件夹\n")
        f.write("3. 设置输出Excel文件名\n")
        f.write("4. 点击开始转换\n\n")
        f.write("命令行使用:\n")
        f.write("文件夹图片转Excel工具.exe <文件夹路径> -o <输出文件名>\n\n")
        f.write("依赖包:\n")
        with open('requirements.txt', 'r', encoding='utf-8') as req_file:
            f.write(req_file.read())
    
    print(f"✓ 构建信息文件已创建: {info_file}")


def main():
    """主函数"""
    print("文件夹图片转Excel工具 - exe构建脚本")
    print("="*50)
    
    # 检查主程序文件是否存在
    main_file = Path("folder_to_excel.py")
    if not main_file.exists():
        print(f"✗ 主程序文件不存在: {main_file}")
        return False
    
    # 创建图标
    create_icon()
    
    # 安装依赖
    if not install_dependencies():
        return False
    
    # 构建exe
    if not build_exe():
        return False
    
    # 创建构建信息
    create_build_info()
    
    print("\n" + "="*50)
    print("✓ 构建完成！")
    print(f"✓ exe文件位置: {Path('dist').absolute()}")
    print("✓ 可以将dist文件夹中的exe文件分发给其他用户使用")
    
    return True


if __name__ == "__main__":
    success = main()
    if not success:
        sys.exit(1) 