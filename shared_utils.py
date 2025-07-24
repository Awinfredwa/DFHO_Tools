#!/usr/bin/env python3
"""
共享工具函数模块
===============

此模块包含可在不同脚本间重复使用的通用函数，
特别是PowerPoint图片替换和图片处理相关的功能。
"""

import os
import sys
from PIL import Image
from pptx import Presentation
from pathlib import Path

# 支持的图片格式
SUPPORTED_EXTENSIONS = [".jpg", ".jpeg", ".png", ".bmp", ".tiff", ".webp"]

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

def print_separator():
    """打印视觉分隔线以提高可读性"""
    print("=" * 60)

def replace_images_in_presentation_simple(pptx_path, image_path, image_name=None):
    """
    替换PowerPoint演示文稿中的图片（简化版本）
    替换第一个找到的图片形状
    
    Args:
        pptx_path: PowerPoint文件路径
        image_path: 要插入的图片路径
        image_name: 图片名称（用于文本替换，可选）
    
    Returns:
        bool: 是否成功替换
    """
    try:
        prs = Presentation(pptx_path)
        replacements_made = 0
        text_replacements_made = 0
        
        print(f"  🔍 正在查找要替换的图片形状...")
        
        # 处理图片替换 - 查找所有图片形状并替换第一个找到的
        for slide_idx, slide in enumerate(prs.slides):
            print(f"  📄 正在处理幻灯片 {slide_idx + 1}")
            
            # 查找第一个图片形状并替换它
            for shape in slide.shapes:
                if shape.shape_type == 13:  # 13表示图片类型
                    print(f"    🔄 正在替换图片形状为{image_path.name}")
                    
                    # 保存原始位置、大小和索引
                    left = shape.left
                    top = shape.top
                    width = shape.width
                    height = shape.height
                    shape_index = slide.shapes.index(shape)
                    
                    # 移除原始形状
                    sp = shape._element
                    sp.getparent().remove(sp)
                    
                    # 添加新图片并保持原始位置和大小
                    new_pic = slide.shapes.add_picture(str(image_path), left, top, width, height)
                    
                    # 确保新图片具有与原始图片相同的属性
                    new_pic.left = left
                    new_pic.top = top
                    new_pic.width = width
                    new_pic.height = height
                    
                    # 正确处理XML元素顺序
                    new_pic_element = new_pic._element
                    slide.shapes._spTree.remove(new_pic_element)
                    
                    # 确保插入索引不会破坏spTree结构
                    spTree = slide.shapes._spTree
                    group_props_count = 0
                    
                    # 计算组属性元素的数量
                    for child in spTree:
                        if child.tag.endswith('}nvGrpSpPr') or child.tag.endswith('}grpSpPr'):
                            group_props_count += 1
                        else:
                            break
                    
                    correct_insert_index = max(group_props_count, shape_index + group_props_count)
                    
                    if correct_insert_index > len(spTree):
                        spTree.append(new_pic_element)
                    else:
                        spTree.insert(correct_insert_index, new_pic_element)
                    
                    replacements_made += 1
                    print(f"    ✅ 成功替换图片形状")
                    break  # 只替换第一个找到的图片
            
            if replacements_made > 0:
                break  # 只处理第一张幻灯片
        
        # 处理文本替换（如果提供了图片名称）
        if image_name:
            print(f"  🔍 正在查找文本框进行替换...")
            
            for slide_idx, slide in enumerate(prs.slides):
                # 查找文本框并替换内容
                for shape in slide.shapes:
                    if hasattr(shape, 'text_frame') and shape.text_frame:
                        # 检查是否有文本内容
                        if shape.text_frame.text.strip():
                            print(f"    📝 正在替换文本为'{image_name}'")
                            
                            # 保存原始格式并替换文本
                            text_frame = shape.text_frame
                            
                            # 如果文本框有内容，保持第一段的格式
                            if text_frame.paragraphs:
                                # 获取第一段
                                first_paragraph = text_frame.paragraphs[0]
                                
                                # 保存原始段落格式
                                original_alignment = first_paragraph.alignment if hasattr(first_paragraph, 'alignment') else None
                                
                                # 如果第一段有运行（runs），保存第一个运行的格式
                                original_font = None
                                if first_paragraph.runs:
                                    first_run = first_paragraph.runs[0]
                                    original_font = {
                                        'name': first_run.font.name,
                                        'size': first_run.font.size,
                                        'bold': first_run.font.bold,
                                        'italic': first_run.font.italic,
                                        'color': first_run.font.color.rgb if hasattr(first_run.font.color, 'rgb') else None
                                    }
                                
                                # 清除现有内容
                                text_frame.clear()
                                
                                # 添加新段落
                                new_paragraph = text_frame.paragraphs[0]
                                
                                # 恢复段落格式
                                if original_alignment is not None:
                                    new_paragraph.alignment = original_alignment
                                
                                # 添加新文本运行
                                new_run = new_paragraph.add_run()
                                new_run.text = image_name
                                
                                # 恢复字体格式
                                if original_font:
                                    if original_font['name']:
                                        new_run.font.name = original_font['name']
                                    if original_font['size']:
                                        new_run.font.size = original_font['size']
                                    if original_font['bold'] is not None:
                                        new_run.font.bold = original_font['bold']
                                    if original_font['italic'] is not None:
                                        new_run.font.italic = original_font['italic']
                                    if original_font['color']:
                                        new_run.font.color.rgb = original_font['color']
                                
                                text_replacements_made += 1
                                print(f"    ✅ 成功更新文本框（已保持格式）")
                                break  # 只替换第一个找到的文本框
        
        # 保存演示文稿
        prs.save(pptx_path)
        print(f"  📊 替换统计:")
        print(f"     - 图片替换: {replacements_made} 个")
        print(f"     - 文本替换: {text_replacements_made} 个")
        
        return replacements_made > 0
        
    except Exception as e:
        print(f"  ❌ 更新演示文稿时出错：{e}")
        return False

def replace_images_in_presentation_advanced(pptx_path, image_files, folder_name=None):
    """
    替换PowerPoint演示文稿中的图片（高级版本）
    支持多个图片文件和特定的形状名称映射
    
    Args:
        pptx_path: PowerPoint文件路径
        image_files: 图片文件列表
        folder_name: 文件夹名称（用于文本替换，可选）
    
    Returns:
        bool: 是否成功替换
    """
    try:
        prs = Presentation(pptx_path)
        replacements_made = 0
        text_replacements_made = 0
        
        # 创建形状名称到图片文件的映射
        image_map = {}
        additional_image = None
        
        # 处理每个输入文件并映射到对应的形状名称
        for img_file in image_files:
            # 获取基本文件名（不带扩展名和_temp后缀）
            base_name = img_file.stem
            if base_name.endswith('_temp'):
                base_name = base_name[:-5]  # 移除_temp后缀
            
            # 根据文件名映射到对应的形状名称
            if base_name == '1':
                image_map['1'] = img_file
            elif base_name == '2':
                image_map['2'] = img_file
            elif base_name == '3':
                image_map['3'] = img_file
            else:
                # 保存额外图片
                additional_image = img_file
        
        # 如果有额外图片，用于替换"4"
        if additional_image:
            image_map['4'] = additional_image
        
        print(f"  🔍 查找要替换的形状：{list(image_map.keys())}")
        
        # 处理图片替换
        for slide_idx, slide in enumerate(prs.slides):
            print(f"  📄 正在处理幻灯片 {slide_idx + 1}")
            
            for shape_name, new_image_path in image_map.items():
                # 检查图片是否存在
                if not new_image_path.exists():
                    print(f"    ❌ 未找到图片：{new_image_path}")
                    continue
                
                # 按名称查找形状
                for shape in slide.shapes:
                    if hasattr(shape, 'name') and shape.name == shape_name and shape.shape_type == 13:  # 13表示图片类型
                        print(f"    🔄 正在替换形状'{shape_name}'为{new_image_path.name}")
                        
                        # 保存原始位置、大小和索引
                        left = shape.left
                        top = shape.top
                        width = shape.width
                        height = shape.height
                        shape_index = slide.shapes.index(shape)
                        
                        # 移除原始形状
                        sp = shape._element
                        sp.getparent().remove(sp)
                        
                        # 添加新图片并保持原始位置和大小
                        new_pic = slide.shapes.add_picture(str(new_image_path), left, top, width, height)
                        
                        # 确保新图片具有与原始图片相同的属性
                        new_pic.left = left
                        new_pic.top = top
                        new_pic.width = width
                        new_pic.height = height
                        
                        # 正确处理XML元素顺序
                        new_pic_element = new_pic._element
                        slide.shapes._spTree.remove(new_pic_element)
                        
                        # 确保插入索引不会破坏spTree结构
                        spTree = slide.shapes._spTree
                        group_props_count = 0
                        
                        # 计算组属性元素的数量
                        for child in spTree:
                            if child.tag.endswith('}nvGrpSpPr') or child.tag.endswith('}grpSpPr'):
                                group_props_count += 1
                            else:
                                break
                        
                        correct_insert_index = max(group_props_count, shape_index + group_props_count)
                        
                        if correct_insert_index > len(spTree):
                            spTree.append(new_pic_element)
                        else:
                            spTree.insert(correct_insert_index, new_pic_element)
                        
                        replacements_made += 1
                        break
        
        # 处理文本替换
        if additional_image or folder_name:
            text_to_use = folder_name if folder_name else (additional_image.stem if additional_image else "")
            
            if text_to_use:
                print(f"  🔍 正在查找文本框进行替换...")
                
                # 定义要替换的文本框名称列表
                text_boxes_to_replace = ['name']
                
                for slide_idx, slide in enumerate(prs.slides):
                    # 查找并替换所有指定的文本框
                    for text_box_name in text_boxes_to_replace:
                        for shape in slide.shapes:
                            if hasattr(shape, 'name') and shape.name == text_box_name:
                                # 检查是否是文本框
                                if hasattr(shape, 'text_frame') and shape.text_frame:
                                    print(f"    📝 正在替换文本框'{shape.name}'为'{text_to_use}'")
                                    
                                    # 使用与simple版本相同的文本替换逻辑
                                    text_frame = shape.text_frame
                                    
                                    if text_frame.paragraphs:
                                        first_paragraph = text_frame.paragraphs[0]
                                        original_alignment = first_paragraph.alignment if hasattr(first_paragraph, 'alignment') else None
                                        
                                        original_font = None
                                        if first_paragraph.runs:
                                            first_run = first_paragraph.runs[0]
                                            original_font = {
                                                'name': first_run.font.name,
                                                'size': first_run.font.size,
                                                'bold': first_run.font.bold,
                                                'italic': first_run.font.italic,
                                                'color': first_run.font.color.rgb if hasattr(first_run.font.color, 'rgb') else None
                                            }
                                        
                                        text_frame.clear()
                                        new_paragraph = text_frame.paragraphs[0]
                                        
                                        if original_alignment is not None:
                                            new_paragraph.alignment = original_alignment
                                        
                                        new_run = new_paragraph.add_run()
                                        new_run.text = text_to_use
                                        
                                        if original_font:
                                            if original_font['name']:
                                                new_run.font.name = original_font['name']
                                            if original_font['size']:
                                                new_run.font.size = original_font['size']
                                            if original_font['bold'] is not None:
                                                new_run.font.bold = original_font['bold']
                                            if original_font['italic'] is not None:
                                                new_run.font.italic = original_font['italic']
                                            if original_font['color']:
                                                new_run.font.color.rgb = original_font['color']
                                        
                                        text_replacements_made += 1
                                        print(f"    ✅ 成功更新文本框'{text_box_name}'（已保持格式）")
                                        break
        
        # 保存演示文稿
        prs.save(pptx_path)
        print(f"  📊 替换统计:")
        print(f"     - 图片替换: {replacements_made} 个")
        print(f"     - 文本替换: {text_replacements_made} 个")
        
        return (replacements_made > 0) or (text_replacements_made > 0)
        
    except Exception as e:
        print(f"  ❌ 更新演示文稿时出错：{e}")
        return False

def create_unfolded_image_four_same(image_path, output_path):
    """
    从单张图片创建完整的展开图片（四面相同）
    
    Args:
        image_path: 输入图片路径
        output_path: 输出图片路径
    
    Returns:
        bool: 处理是否成功
    """
    try:
        with Image.open(image_path) as original:
            print(f"  📐 原图尺寸：{original.size[0]}x{original.size[1]}像素")
            
            orig_width, orig_height = original.size
            new_width = orig_width * 2
            new_height = orig_height * 2
            new_image = Image.new('RGB', (new_width, new_height), 'white')
            
            print(f"  🔧 创建展开画布：{new_width}x{new_height}像素")
            
            # 四个象限都放置相同的原图
            new_image.paste(original, (0, 0))  # 左上角
            new_image.paste(original, (orig_width, 0))  # 右上角
            new_image.paste(original, (0, orig_height))  # 左下角
            new_image.paste(original, (orig_width, orig_height))  # 右下角
            
            new_image.save(output_path, quality=95, optimize=True)
            print(f"  💾 已保存展开图：{output_path.name}")
            
            return True
            
    except Exception as e:
        print(f"  ❌ 处理图片时出错：{e}")
        return False

def create_unfolded_image_with_rotation(image_path, output_path):
    """
    从单张图片创建完整的展开图片（带旋转的版本）
    
    Args:
        image_path: 输入图片路径
        output_path: 输出图片路径
    
    Returns:
        bool: 处理是否成功
    """
    try:
        with Image.open(image_path) as original:
            print(f"  📐 原图尺寸：{original.size[0]}x{original.size[1]}像素")
            
            orig_width, orig_height = original.size
            new_width = orig_width * 2
            new_height = orig_height * 2
            new_image = Image.new('RGB', (new_width, new_height), 'white')
            
            print(f"  🔧 创建展开画布：{new_width}x{new_height}像素")
            
            # 右下角：原图
            new_image.paste(original, (orig_width, orig_height))
            # 左下角：原图复制
            new_image.paste(original, (0, orig_height))
            
            # 创建下半部分并旋转180度作为上半部分
            bottom_half = Image.new('RGB', (new_width, orig_height), 'white')
            bottom_half.paste(original, (0, 0))
            bottom_half.paste(original, (orig_width, 0))
            top_half = bottom_half.rotate(180)
            new_image.paste(top_half, (0, 0))
            
            new_image.save(output_path, quality=95, optimize=True)
            print(f"  💾 已保存展开图：{output_path.name}")
            
            return True
            
    except Exception as e:
        print(f"  ❌ 处理图片时出错：{e}")
        return False 