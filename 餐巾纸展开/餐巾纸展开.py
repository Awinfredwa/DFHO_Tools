#!/usr/bin/env python3
"""
餐巾纸展开图片处理脚本
====================

此脚本自动完成以下任务：
1. 从"展开原图（记得在处理完成后清空）"文件夹读取所有图片
2. 将单张图片展开为完整的4倍大小餐巾纸
3. 新的展开逻辑：
   - 右下角：原图
   - 左下角：原图的复制（不变）
   - 上半部分（左上角+右上角）：下半部分（左下角+右下角）旋转180度

简易设置：
1. 将此脚本放在您的工作文件夹中
2. 在"展开原图（记得在处理完成后清空）"文件夹中放置图片文件（支持JPG、PNG格式）
3. 处理后的图片将保存到"展开成果"文件夹
4. 记得在处理完成后清空原图文件夹

专为非技术用户设计 - 只需点击运行！
"""

import os
import sys
import shutil
from PIL import Image
from pathlib import Path

# 支持的图片格式
SUPPORTED_EXTENSIONS = [".jpg", ".jpeg", ".png", ".bmp", ".tiff", ".webp"]

# 文件夹名称
INPUT_FOLDER = "展开原图（记得在处理完成后清空）"
OUTPUT_FOLDER = "展开成果"

def print_separator():
    """打印视觉分隔线以提高可读性"""
    print("=" * 60)

def get_script_directory():
    """获取脚本所在目录"""
    if getattr(sys, 'frozen', False):
        # 如果作为exe运行（使用PyInstaller编译）
        return Path(sys.executable).parent
    else:
        # 如果作为Python脚本运行
        return Path(__file__).parent

def find_images_in_folder(folder_path):
    """查找文件夹中的所有图片文件"""
    image_files = []
    
    if not folder_path.exists():
        return image_files
    
    for ext in SUPPORTED_EXTENSIONS:
        image_files.extend(list(folder_path.glob(f"*{ext}")))
        image_files.extend(list(folder_path.glob(f"*{ext.upper()}")))
    
    return image_files

def create_unfolded_napkin(image_path, output_path):
    """
    从单张图片创建完整的展开餐巾纸
    新逻辑：
    - 右下角：原图
    - 左下角：原图复制（不变）
    - 上半部分：下半部分旋转180度
    
    Args:
        image_path: 输入图片路径
        output_path: 输出图片路径
    
    Returns:
        bool: 处理是否成功
    """
    try:
        # 打开原始图片
        with Image.open(image_path) as original:
            print(f"  📐 原图尺寸：{original.size[0]}x{original.size[1]}像素")
            
            # 获取原图尺寸
            orig_width, orig_height = original.size
            
            # 创建4倍大小的新画布
            new_width = orig_width * 2
            new_height = orig_height * 2
            new_image = Image.new('RGB', (new_width, new_height), 'white')
            
            print(f"  🔧 创建展开画布：{new_width}x{new_height}像素")
            
            # 1. 右下角：放置原图
            new_image.paste(original, (orig_width, orig_height))
            print(f"  📍 右下角：原图")
            
            # 2. 左下角：原图的复制（不变）
            new_image.paste(original, (0, orig_height))
            print(f"  📍 左下角：原图复制")
            
            # 3. 创建下半部分（左下角+右下角的组合）
            bottom_half = Image.new('RGB', (new_width, orig_height), 'white')
            bottom_half.paste(original, (0, 0))      # 左下角
            bottom_half.paste(original, (orig_width, 0))  # 右下角
            
            # 4. 将下半部分旋转180度作为上半部分
            top_half = bottom_half.rotate(180)
            new_image.paste(top_half, (0, 0))
            print(f"  📍 上半部分：下半部分旋转180度")
            
            # 保存展开的餐巾纸图片
            new_image.save(output_path, quality=95, optimize=True)
            print(f"  💾 已保存展开图：{output_path.name}")
            
            return True
            
    except Exception as e:
        print(f"  ❌ 处理图片时出错：{e}")
        return False

def process_image(image_path, output_folder):
    """处理单张图片"""
    print(f"\n📷 正在处理图片：{image_path.name}")
    
    # 生成输出文件名
    base_name = image_path.stem
    extension = image_path.suffix
    output_name = f"{base_name}_unfolded{extension}"
    output_path = output_folder / output_name
    
    # 创建展开的餐巾纸图片
    print(f"  🔧 正在展开餐巾纸...")
    success = create_unfolded_napkin(image_path, output_path)
    
    if success:
        print(f"  🎉 成功处理图片：{image_path.name}")
        return True
    else:
        print(f"  ❌ 处理失败：{image_path.name}")
        return False

def main():
    """主函数，协调整个处理过程"""
    try:
        # 自动检测工作目录
        working_dir = get_script_directory()
        input_folder = working_dir / INPUT_FOLDER
        output_folder = working_dir / OUTPUT_FOLDER
        
        print_separator()
        print("🧾 餐巾纸展开处理器")
        print_separator()
        print("此脚本将自动处理输入文件夹中的所有图片，")
        print("并将图片展开为完整的4倍大小餐巾纸图案。")
        print(f"工作目录：{working_dir}")
        print(f"输入文件夹：{INPUT_FOLDER}")
        print(f"输出文件夹：{OUTPUT_FOLDER}")
        print()
        print("📋 展开模式说明：")
        print("  新的展开逻辑：")
        print("    ┌─────────┬─────────┐")
        print("    │ 左上角  │ 右上角  │")
        print("    │    下半部分旋转180度   │")
        print("    ├─────────┼─────────┤")
        print("    │ 左下角  │ 右下角  │")
        print("    │ 原图复制 │  原图   │")
        print("    └─────────┴─────────┘")
        print()
        print(f"📸 支持的图片格式：{', '.join(SUPPORTED_EXTENSIONS)}")
        print()
        
        # 创建输入文件夹（如果不存在）
        if not input_folder.exists():
            input_folder.mkdir(exist_ok=True)
            print(f"✅ 已创建输入文件夹：{INPUT_FOLDER}")
            print("请将要处理的图片放入此文件夹中。")
            input("按回车键退出...")
            return
        
        # 创建输出文件夹（如果不存在）
        output_folder.mkdir(exist_ok=True)
        
        print_separator()
        print("🔍 扫描输入文件夹")
        print_separator()
        
        # 查找输入文件夹中的所有图片
        image_files = find_images_in_folder(input_folder)
        
        if not image_files:
            print(f"❌ 在输入文件夹「{INPUT_FOLDER}」中未找到图片文件。")
            print("请确保：")
            print("  1. 图片文件放在正确的输入文件夹中")
            print(f"  2. 图片格式为：{', '.join(SUPPORTED_EXTENSIONS)}")
            input("按回车键退出...")
            return
        
        print(f"✅ 找到{len(image_files)}张图片")
        for img_file in image_files:
            print(f"  📷 {img_file.name}")
        
        # 确认开始处理
        print()
        response = input("🚀 准备开始处理？（按回车继续，输入'n'取消）：").lower()
        if response == 'n':
            print("操作已取消。")
            return
        
        print()
        print_separator()
        print("🚀 开始处理")
        print_separator()
        
        # 处理每张图片
        successful = 0
        for image_path in image_files:
            if process_image(image_path, output_folder):
                successful += 1
        
        # 最终总结
        print()
        print_separator()
        print("✨ 处理完成")
        print_separator()
        print(f"成功处理：{successful}/{len(image_files)}张图片")
        
        if successful > 0:
            print(f"\n🎉 您的餐巾纸展开图已准备就绪！")
            print(f"处理结果保存在：{OUTPUT_FOLDER}")
            print("每张处理后的图片命名格式：[原文件名]_unfolded[扩展名]")
            print()
            print("📝 重要提醒：")
            print(f"   处理完成后，请记得清空「{INPUT_FOLDER}」文件夹")
            print("   以便下次使用。")
        else:
            print("\n⚠️  未成功处理任何图片。")
            print("请检查您的图片文件是否正确。")
        
        # 保持窗口打开
        print()
        input("按回车键退出...")
        
    except KeyboardInterrupt:
        print("\n\n⏹️  用户取消操作。")
        input("按回车键退出...")
    except Exception as e:
        print(f"\n❌ 意外错误：{e}")
        print("请检查您的文件并重试。")
        input("按回车键退出...")

if __name__ == "__main__":
    main()

