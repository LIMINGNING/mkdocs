#!/usr/bin/env python3
"""
PDF to PNG Converter
无损将PDF转换为PNG格式
支持高分辨率和多页PDF处理
"""

import os
import sys
import argparse
from pathlib import Path

try:
    import fitz  # PyMuPDF
    PYMUPDF_AVAILABLE = True
except ImportError:
    PYMUPDF_AVAILABLE = False

try:
    from pdf2image import convert_from_path
    from PIL import Image
    PDF2IMAGE_AVAILABLE = True
except ImportError:
    PDF2IMAGE_AVAILABLE = False


def convert_pdf_with_pymupdf(pdf_path, output_dir, dpi=300):
    """
    使用PyMuPDF (fitz) 转换PDF为PNG
    这是最推荐的方法，速度快且质量高
    """
    print(f"使用PyMuPDF转换: {pdf_path}")
    
    # 打开PDF文档
    pdf_document = fitz.open(pdf_path)
    
    # 获取输出文件名前缀
    pdf_name = Path(pdf_path).stem
    
    converted_files = []
    
    for page_num in range(len(pdf_document)):
        # 获取页面
        page = pdf_document.load_page(page_num)
        
        # 设置变换矩阵以获得高分辨率
        # dpi/72.0 是因为PDF默认是72 DPI
        mat = fitz.Matrix(dpi/72.0, dpi/72.0)
        
        # 渲染页面为图像
        pix = page.get_pixmap(matrix=mat, alpha=False)
        
        # 生成输出文件名
        if len(pdf_document) == 1:
            output_file = os.path.join(output_dir, f"{pdf_name}.png")
        else:
            output_file = os.path.join(output_dir, f"{pdf_name}_page_{page_num + 1:03d}.png")
        
        # 保存为PNG
        pix.save(output_file)
        converted_files.append(output_file)
        
        print(f"已转换页面 {page_num + 1}/{len(pdf_document)}: {output_file}")
    
    pdf_document.close()
    return converted_files


def convert_pdf_with_pdf2image(pdf_path, output_dir, dpi=300):
    """
    使用pdf2image转换PDF为PNG
    需要安装poppler-utils
    """
    print(f"使用pdf2image转换: {pdf_path}")
    
    # 转换PDF为PIL图像列表
    images = convert_from_path(pdf_path, dpi=dpi, fmt='png')
    
    # 获取输出文件名前缀
    pdf_name = Path(pdf_path).stem
    
    converted_files = []
    
    for i, image in enumerate(images):
        # 生成输出文件名
        if len(images) == 1:
            output_file = os.path.join(output_dir, f"{pdf_name}.png")
        else:
            output_file = os.path.join(output_dir, f"{pdf_name}_page_{i + 1:03d}.png")
        
        # 保存为PNG
        image.save(output_file, "PNG", optimize=False, compress_level=0)
        converted_files.append(output_file)
        
        print(f"已转换页面 {i + 1}/{len(images)}: {output_file}")
    
    return converted_files


def install_dependencies():
    """安装必要的依赖"""
    print("正在安装依赖...")
    
    # 尝试安装PyMuPDF (推荐)
    try:
        import subprocess
        subprocess.check_call([sys.executable, "-m", "pip", "install", "PyMuPDF"])
        print("PyMuPDF 安装成功!")
        return True
    except subprocess.CalledProcessError:
        print("PyMuPDF 安装失败，尝试安装pdf2image...")
        
    # 尝试安装pdf2image
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pdf2image", "Pillow"])
        print("pdf2image 安装成功!")
        print("注意: pdf2image 还需要系统安装 poppler-utils")
        print("Ubuntu/Debian: sudo apt-get install poppler-utils")
        print("CentOS/RHEL: sudo yum install poppler-utils")
        print("macOS: brew install poppler")
        return True
    except subprocess.CalledProcessError:
        print("依赖安装失败!")
        return False


def main():
    parser = argparse.ArgumentParser(description="无损PDF转PNG转换器")
    parser.add_argument("input", nargs="?", help="输入PDF文件路径")
    parser.add_argument("-o", "--output", help="输出目录路径 (默认: 与输入文件同目录)")
    parser.add_argument("-d", "--dpi", type=int, default=300, help="输出DPI (默认: 300)")
    parser.add_argument("--install-deps", action="store_true", help="安装必要的依赖")
    parser.add_argument("--auto", action="store_true", help="自动处理当前目录下的PDF文件")
    
    args = parser.parse_args()
    
    # 安装依赖
    if args.install_deps:
        success = install_dependencies()
        if not success:
            sys.exit(1)
        return
    
    # 检查是否有可用的转换库
    if not PYMUPDF_AVAILABLE and not PDF2IMAGE_AVAILABLE:
        print("错误: 没有找到可用的PDF转换库!")
        print("请运行以下命令安装依赖:")
        print("python pdf_to_png.py --install-deps")
        print("\n或手动安装:")
        print("pip install PyMuPDF  # 推荐")
        print("# 或者")
        print("pip install pdf2image Pillow")
        sys.exit(1)
    
    # 自动模式：处理当前目录下的PDF文件
    if args.auto:
        pdf_files = list(Path(".").rglob("*.pdf"))
        if not pdf_files:
            print("当前目录下没有找到PDF文件")
            return
        
        for pdf_file in pdf_files:
            print(f"\n处理文件: {pdf_file}")
            output_dir = pdf_file.parent
            
            try:
                if PYMUPDF_AVAILABLE:
                    converted_files = convert_pdf_with_pymupdf(str(pdf_file), str(output_dir), args.dpi)
                else:
                    converted_files = convert_pdf_with_pdf2image(str(pdf_file), str(output_dir), args.dpi)
                
                print(f"成功转换 {len(converted_files)} 个文件")
                
            except Exception as e:
                print(f"转换失败: {e}")
        
        return
    
    # 处理指定的PDF文件
    if not args.input:
        print("请指定输入PDF文件或使用 --auto 参数")
        parser.print_help()
        sys.exit(1)
    
    pdf_path = args.input
    
    # 检查输入文件
    if not os.path.exists(pdf_path):
        print(f"错误: 文件不存在 - {pdf_path}")
        sys.exit(1)
    
    if not pdf_path.lower().endswith('.pdf'):
        print(f"错误: 不是PDF文件 - {pdf_path}")
        sys.exit(1)
    
    # 设置输出目录
    if args.output:
        output_dir = args.output
    else:
        output_dir = os.path.dirname(pdf_path) or "."
    
    # 创建输出目录
    os.makedirs(output_dir, exist_ok=True)
    
    try:
        # 选择转换方法
        if PYMUPDF_AVAILABLE:
            converted_files = convert_pdf_with_pymupdf(pdf_path, output_dir, args.dpi)
        else:
            converted_files = convert_pdf_with_pdf2image(pdf_path, output_dir, args.dpi)
        
        print(f"\n转换完成! 共生成 {len(converted_files)} 个PNG文件:")
        for file in converted_files:
            file_size = os.path.getsize(file) / (1024 * 1024)  # MB
            print(f"  {file} ({file_size:.2f} MB)")
            
    except Exception as e:
        print(f"转换失败: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()