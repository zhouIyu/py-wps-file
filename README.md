# 图片处理工具集

这是一个包含多个图片处理功能的Python工具集。

## 工具列表

### 1. PDF图片提取工具 (`extract_pdf_images.py`)
从PDF文件中提取所有图片并保存到指定目录。

### 2. 文件夹图片到Word文档工具 (`folder_to_word.py`)
将文件夹中的二级目录作为标题，并将目录内的图片插入到Word文档中。

### 3. 文件夹图片到Excel工具 (`folder_to_excel.py`)
将文件夹中的二级目录名称作为Excel工作表（sheet），并将目录内的图片插入到对应的工作表中。

### 4. 文件夹图片到PDF文档工具 (`folder_to_pdf.py`) 🆕
将文件夹中的二级目录作为标题，并将目录内的图片插入到PDF文档中。

## 安装依赖

### 方法1：使用虚拟环境（推荐）
```bash
# 创建虚拟环境
python3 -m venv venv

# 激活虚拟环境
source venv/bin/activate  # macOS/Linux
# 或
venv\Scripts\activate     # Windows

# 安装依赖
pip install -r requirements.txt
```

### 方法2：直接安装
```bash
pip install -r requirements.txt
# 或者如果遇到权限问题
pip install --user -r requirements.txt
```

---

## 1. PDF图片提取工具

### 功能特点
- ✅ 支持从PDF的所有页面提取图片
- ✅ 自动创建输出目录
- ✅ 保持原始图片格式（PNG, JPEG等）
- ✅ 自动命名：`page_X_img_Y.extension`
- ✅ 详细的进度显示和错误处理

### 使用方法

```bash
# 基本用法
python extract_pdf_images.py your_file.pdf

# 指定输出目录
python extract_pdf_images.py your_file.pdf -o output_folder
```

### 输出示例
提取的图片将按以下格式命名：
- `page_1_img_1.png`
- `page_1_img_2.jpg`
- `page_2_img_1.png`

---

## 2. 文件夹图片到Word文档工具 🆕

### 功能特点
- ✅ 遍历文件夹中的所有二级目录
- ✅ 将目录名作为Word文档中的标题
- ✅ 自动插入目录内的所有图片
- ✅ 支持多种图片格式（jpg, jpeg, png, gif, bmp, tiff, webp）
- ✅ 自动调整图片尺寸以适合Word文档
- ✅ 生成美观的文档布局，包含图片说明和统计信息

### 使用方法

#### 命令行方式：
```bash
# 基本用法
python folder_to_word.py /path/to/your/folder

# 指定输出文件名
python folder_to_word.py /path/to/your/folder -o "我的图片文档.docx"
```

#### 代码调用方式：
```python
from folder_to_word import create_word_document

# 创建Word文档
success = create_word_document("/path/to/your/folder", "输出文档.docx")
if success:
    print("文档创建成功！")
```

#### 示例使用：
```bash
# 运行示例脚本
python folder_to_word_example.py
```

### 文件夹结构示例

假设您的文件夹结构如下：
```
我的图片文件夹/
├── 第一章/
│   ├── 图片1.jpg
│   ├── 图片2.png
│   └── 图片3.gif
├── 第二章/
│   ├── photo1.jpeg
│   └── photo2.bmp
└── 第三章/
    ├── image1.tiff
    ├── image2.webp
    └── image3.jpg
```

生成的Word文档将包含：
- **文档标题**: "我的图片文件夹 - 图片文档"
- **第一章**（标题）
  - 图片数量：3
  - 图片 1: 图片1.jpg
  - 图片 2: 图片2.png
  - 图片 3: 图片3.gif
- **第二章**（标题）
  - 图片数量：2
  - 图片 1: photo1.jpeg
  - 图片 2: photo2.bmp
- **第三章**（标题）
  - 图片数量：3
  - 图片 1: image1.tiff
  - 图片 2: image2.webp
  - 图片 3: image3.jpg
- **统计信息**
  - 处理的目录数量：3
  - 插入的图片数量：8

### 支持的图片格式
- JPG/JPEG
- PNG  
- GIF
- BMP
- TIFF
- WebP

### 注意事项
1. 确保目标文件夹存在且有读取权限
2. 图片文件会自动调整尺寸以适合Word文档
3. 目录和图片会按名称排序
4. 每个章节之间会自动分页
5. 如果某个目录没有图片，会跳过该目录

---

## 4. 文件夹图片到PDF文档工具 🆕

### 功能特点
- ✅ 遍历文件夹中的所有二级目录
- ✅ 将目录名作为PDF文档中的标题
- ✅ 自动插入目录内的所有图片
- ✅ 支持多种图片格式（jpg, jpeg, png, gif, bmp, tiff, webp）
- ✅ 自动调整图片尺寸以适合PDF页面
- ✅ 生成美观的PDF布局，包含封面页和目录页
- ✅ 自动分页和图片居中布局
- ✅ 详细的日志记录和错误处理

### 使用方法

#### 命令行方式：
```bash
# 基本用法
python folder_to_pdf.py /path/to/your/folder

# 指定输出文件名
python folder_to_pdf.py /path/to/your/folder -o "我的图片文档.pdf"
```

#### 代码调用方式：
```python
from folder_to_pdf import create_pdf_document

# 创建PDF文档
success = create_pdf_document("/path/to/your/folder", "输出文档.pdf")
if success:
    print("PDF文档创建成功！")
```

#### 示例使用：
```bash
# 运行示例脚本
python folder_to_pdf_example.py
```

### 文件夹结构示例

假设您的文件夹结构如下：
```
我的图片文件夹/
├── 第一章/
│   ├── 图片1.jpg
│   ├── 图片2.png
│   └── 图片3.gif
├── 第二章/
│   ├── photo1.jpeg
│   └── photo2.bmp
└── 第三章/
    ├── image1.tiff
    ├── image2.webp
    └── image3.jpg
```

生成的PDF文档将包含：
- **封面页**: "我的图片文件夹 - 图片文档"
- **目录页**: 
  - 1. 第一章 (3张图片)
  - 2. 第二章 (2张图片)  
  - 3. 第三章 (3张图片)
- **内容页面**:
  - 第一章页面：显示3张图片，居中排列
  - 第二章页面：显示2张图片，居中排列
  - 第三章页面：显示3张图片，居中排列
- **统计信息**
  - 处理的目录数量：3
  - 插入的图片数量：8

### 支持的图片格式
- JPG/JPEG
- PNG  
- GIF
- BMP
- TIFF
- WebP

### 注意事项
1. 确保目标文件夹存在且有读取权限
2. 图片文件会自动调整尺寸以适合PDF页面
3. 目录和图片会按名称排序
4. 图片过大时会自动换页
5. 如果某个目录没有图片，会跳过该目录
6. 每张图片下方会显示文件名标注

---

## 技术依赖

- **Python 3.6+**
PyMuPDF>=1.23.0
python-docx>=0.8.11
Pillow>=9.0.0
openpyxl>=3.0.0
pyinstaller>=5.0.0

---

## 3. 文件夹图片到Excel工具 🆕

### 功能特点
- ✅ 遍历文件夹中的所有二级目录
- ✅ 将每个目录名作为Excel工作表（sheet）
- ✅ 自动插入目录内的所有图片到对应工作表中
- ✅ 支持多种图片格式（jpg, jpeg, png, gif, bmp, tiff, webp）
- ✅ 自动调整图片尺寸以适合Excel单元格
- ✅ 美观的GUI界面，支持拖拽操作
- ✅ 实时进度显示和日志记录
- ✅ 支持命令行和图形界面两种使用方式

### 使用方法

#### GUI界面（推荐）:
```bash
# 启动图形界面
python folder_to_excel.py
```

#### 命令行方式：
```bash
# 基本用法
python folder_to_excel.py /path/to/your/folder

# 指定输出文件名
python folder_to_excel.py /path/to/your/folder -o "我的图片表格.xlsx"
```

#### 代码调用方式：
```python
from folder_to_excel import create_excel_document

# 创建Excel文档
success = create_excel_document("/path/to/your/folder", "输出表格.xlsx")
if success:
    print("Excel文档创建成功！")
```

### Excel文档结构示例

假设您的文件夹结构如下：
```
我的图片文件夹/
├── 第一章/
│   ├── 图片1.jpg
│   ├── 图片2.png
│   └── 图片3.gif
├── 第二章/
│   ├── photo1.jpeg
│   └── photo2.bmp
└── 第三章/
    ├── image1.tiff
    ├── image2.webp
    └── image3.jpg
```

生成的Excel文档将包含：
- **工作表1**: "第一章"
  - 标题行：第一章
  - 图片展示：3张图片垂直排列在A列
- **工作表2**: "第二章"
  - 标题行：第二章
  - 图片展示：2张图片垂直排列
- **工作表3**: "第三章"
  - 标题行：第三章
  - 图片展示：3张图片垂直排列

### GUI界面功能
- 📁 文件夹选择：点击"浏览"按钮选择包含图片的文件夹
- 💾 输出设置：可自定义Excel文件名和保存位置
- ⚡ 开始转换：一键启动转换流程
- 📊 实时进度：进度条显示转换进度
- 📝 日志记录：详细的操作日志和错误信息
- 🎯 状态提示：当前操作状态实时显示

---

## 4. 生成可执行文件（exe）

### 功能特点
- ✅ 将Python脚本打包成独立的exe文件
- ✅ 无需安装Python环境即可运行
- ✅ 自动创建应用图标
- ✅ 生成构建信息文档
- ✅ 支持GUI和命令行两种模式

### 构建exe文件

```bash
# 运行构建脚本
python build_exe.py
```

构建完成后，在 `dist` 文件夹中会生成：
- `文件夹图片转Excel工具.exe` - 主程序
- 相关依赖文件（如果需要）

### exe文件使用方法

1. **GUI模式**（双击运行）：
   - 双击 `文件夹图片转Excel工具.exe`
   - 使用图形界面选择文件夹和设置输出

2. **命令行模式**：
   ```cmd
   文件夹图片转Excel工具.exe "C:\path\to\folder" -o "output.xlsx"
   ```

### 分发说明
- exe文件可以在其他Windows系统上直接运行
- 无需安装Python或任何依赖包
- 文件大小约为50-100MB（包含所有依赖）

---

## 注意事项

1. **文件权限**：确保目标文件夹有读取权限
2. **图片格式**：支持常见图片格式，建议使用JPG、PNG
3. **文件大小**：大图片会自动压缩以适合Excel单元格
4. **内存使用**：处理大量图片时可能需要较多内存
5. **Excel限制**：工作表名称不能超过31个字符，会自动截断

---

## 技术依赖

- **Python 3.6+**
- **PyMuPDF (fitz)** - PDF处理
- **python-docx** - Word文档生成
- **openpyxl** - Excel文档处理
- **Pillow (PIL)** - 图片处理
- **tkinter** - GUI界面（Python内置）
- **PyInstaller** - exe文件打包

## 系统要求

- 支持 Windows, macOS, Linux
- 推荐内存：4GB以上（处理大量图片时）
- 硬盘空间：至少100MB（exe文件）

## 开源协议

MIT License 