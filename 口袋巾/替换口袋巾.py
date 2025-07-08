#!/usr/bin/env python3
"""
口袋巾图片处理脚本
================

此脚本自动完成以下任务：
1. 查找文件夹中的图片文件（支持JPG或PNG格式）
2. 在PowerPoint演示文稿中替换指定的图片形状
3. 保持原始图片的尺寸和层级

简易设置：
1. 将此脚本放在您的工作文件夹中（与template.pptx一起）
2. 在子文件夹中整理您的图片文件
3. 每个子文件夹必须包含恰好3张图片文件（JPG或PNG格式，任意名称）
4. 脚本将自动分配图片并根据预设映射替换PowerPoint中的图片形状

专为非技术用户设计 - 只需点击运行！
"""

import os
import sys
import shutil
from PIL import Image
from pptx import Presentation
from pathlib import Path

# 支持的图片格式
SUPPORTED_EXTENSIONS = [".jpg", ".png", ".jpeg"]
# 每个文件夹需要的图片数量
REQUIRED_IMAGE_COUNT = 3

# 图片到PowerPoint形状名称的映射（自动分配）
# 第一张图片 → 5个形状，第二张图片 → 2个形状，第三张图片 → 1个形状
IMAGE_SHAPE_MAPPING = [
    ["图片 14", "图片 19", "图片 31", "图片 25", "图片 17"],  # 第一张图片 (5个形状)
    ["图片 43", "图片 18"],                                    # 第二张图片 (2个形状)  
    ["图片 47"]                                              # 第三张图片 (1个形状)
]

# 图片名称到PowerPoint文本框的映射（自动分配）
# 图片名称（去掉扩展名）将替换这些文本框中的文本
TEXT_REPLACEMENT_MAPPING = [
    ["文本框 26", "文本框 36", "文本框 24", "文本框 14"],      # 第一张图片名称 (4个文本框)
    ["文本框 27", "文本框 11"],                               # 第二张图片名称 (2个文本框)
    ["文本框 32"]                                           # 第三张图片名称 (1个文本框)
]

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

def find_template_pptx(working_dir):
    """自动查找模板PPTX文件"""
    print("🔍 正在查找PowerPoint模板文件...")
    pptx_files = list(working_dir.glob("*.pptx"))
    
    if not pptx_files:
        print("❌ 在工作文件夹中未找到PowerPoint（.pptx）文件。")
        print("请确保您的模板PowerPoint文件与此脚本在同一文件夹中。")
        input("按回车键退出...")
        return None
    
    if len(pptx_files) == 1:
        template_pptx = pptx_files[0]
        print(f"✅ 找到模板：{template_pptx.name}")
        return template_pptx
    else:
        print("📋 找到多个PowerPoint文件：")
        for i, pptx in enumerate(pptx_files, 1):
            print(f"  {i}. {pptx.name}")
        
        while True:
            try:
                print()
                choice = input("请输入要使用的模板文件编号（或按回车选择第1个）：").strip()
                if not choice:
                    choice = "1"
                choice = int(choice)
                if 1 <= choice <= len(pptx_files):
                    template_pptx = pptx_files[choice - 1]
                    print(f"✅ 已选择模板：{template_pptx.name}")
                    return template_pptx
                else:
                    print(f"❌ 请输入1到{len(pptx_files)}之间的数字")
            except ValueError:
                print("❌ 请输入有效数字")

def check_required_files(folder_path):
    """检查文件夹中是否存在所需数量的图片文件（支持JPG、PNG、JPEG格式）"""
    # 查找所有支持的图片文件
    image_files = []
    for ext in SUPPORTED_EXTENSIONS:
        image_files.extend(list(folder_path.glob(f"*{ext}")))
        image_files.extend(list(folder_path.glob(f"*{ext.upper()}")))
    
    # 去重并排序以确保一致的处理顺序
    image_files = list(set(image_files))
    image_files.sort(key=lambda x: x.name.lower())
    
    if len(image_files) < REQUIRED_IMAGE_COUNT:
        missing_count = REQUIRED_IMAGE_COUNT - len(image_files)
        return image_files, [f"还需要{missing_count}张图片文件 (支持格式: {', '.join(SUPPORTED_EXTENSIONS)})"]
    elif len(image_files) > REQUIRED_IMAGE_COUNT:
        # 只取前3张图片，警告用户
        used_files = image_files[:REQUIRED_IMAGE_COUNT]
        print(f"  ⚠️  找到{len(image_files)}张图片，将使用前{REQUIRED_IMAGE_COUNT}张：")
        for i, img_file in enumerate(used_files, 1):
            print(f"    {i}. {img_file.name}")
        return used_files, []
    else:
        return image_files, []

def copy_template_pptx(template_path, output_dir, folder_name):
    """创建模板PPTX的副本用于处理"""
    output_name = f"{folder_name}_presentation.pptx"
    output_path = output_dir / output_name
    
    # 处理已存在的文件
    counter = 1
    while output_path.exists():
        output_name = f"{folder_name}_presentation_{counter}.pptx"
        output_path = output_dir / output_name
        counter += 1
    
    try:
        shutil.copy2(template_path, output_path)
        print(f"  📋 已创建演示文稿副本：{output_name}")
        return output_path
    except Exception as e:
        print(f"  ❌ 复制模板失败：{e}")
        return None

def replace_images_in_presentation(pptx_path, image_files):
    """替换PowerPoint演示文稿中的图片和文本"""
    try:
        prs = Presentation(pptx_path)
        replacements_made = 0
        text_replacements_made = 0
        
        # 创建图片文件到形状名称列表的映射
        print(f"  🔍 准备替换以下图片映射：")
        for i, img_file in enumerate(image_files):
            if i < len(IMAGE_SHAPE_MAPPING):
                shape_names = IMAGE_SHAPE_MAPPING[i]
                print(f"    {img_file.name} → {', '.join(shape_names)}")
        
        # 显示文本替换映射
        print(f"  📝 准备替换以下文本映射：")
        for i, img_file in enumerate(image_files):
            if i < len(TEXT_REPLACEMENT_MAPPING):
                text_boxes = TEXT_REPLACEMENT_MAPPING[i]
                image_name = img_file.stem  # 去掉文件扩展名
                print(f"    '{image_name}' → {', '.join(text_boxes)}")
        
        print(f"  🎯 开始处理替换...")
        
        for slide_idx, slide in enumerate(prs.slides):
            print(f"  📄 正在处理幻灯片 {slide_idx + 1}")
            
            # 处理图片替换
            for i, img_file in enumerate(image_files):
                if i >= len(IMAGE_SHAPE_MAPPING):
                    continue
                    
                shape_names = IMAGE_SHAPE_MAPPING[i]
                
                # 检查图片是否存在
                if not img_file.exists():
                    print(f"    ❌ 未找到图片：{img_file}")
                    continue
                
                # 为每个形状名称查找并替换
                for shape_name in shape_names:
                    # 按名称查找形状
                    for shape in slide.shapes:
                        current_shape_name = getattr(shape, 'name', '')
                        if current_shape_name == shape_name:
                            if shape.shape_type == 13:  # 13表示图片类型
                                print(f"    🔄 正在替换形状'{shape_name}'为{img_file.name}")
                                
                                # 保存原始位置、大小和索引
                                left = shape.left
                                top = shape.top
                                width = shape.width
                                height = shape.height
                                shape_index = slide.shapes.index(shape)  # 获取形状的索引
                                
                                # 移除原始形状
                                sp = shape._element  # 获取形状的XML元素
                                sp.getparent().remove(sp)  # 从幻灯片中移除形状
                                
                                # 添加新图片并保持原始位置和大小
                                new_pic = slide.shapes.add_picture(str(img_file), left, top, width, height)
                                
                                # 确保新图片具有与原始图片相同的属性
                                new_pic.left = left
                                new_pic.top = top
                                new_pic.width = width
                                new_pic.height = height
                                
                                # 正确处理XML元素顺序以防止损坏
                                new_pic_element = new_pic._element
                                slide.shapes._spTree.remove(new_pic_element)  # 从当前位置移除
                                
                                # 确保插入索引不会破坏spTree结构
                                # spTree的前两个元素必须是组属性 (nvGrpSpPr 和 grpSpPr)
                                # 所以实际的形状插入索引需要加2
                                spTree = slide.shapes._spTree
                                group_props_count = 0
                                
                                # 计算组属性元素的数量（通常是2个）
                                for child in spTree:
                                    if child.tag.endswith('}nvGrpSpPr') or child.tag.endswith('}grpSpPr'):
                                        group_props_count += 1
                                    else:
                                        break  # 组属性应该在开头，一旦遇到其他元素就停止
                                
                                # 计算正确的插入位置（组属性之后）
                                correct_insert_index = max(group_props_count, shape_index + group_props_count)
                                
                                # 确保不超出范围
                                if correct_insert_index > len(spTree):
                                    spTree.append(new_pic_element)
                                else:
                                    spTree.insert(correct_insert_index, new_pic_element)
                                
                                print(f"    ✅ 成功替换'{shape_name}'")
                                replacements_made += 1
                                break
            
            # 处理文本替换
            for i, img_file in enumerate(image_files):
                if i >= len(TEXT_REPLACEMENT_MAPPING):
                    continue
                    
                text_box_names = TEXT_REPLACEMENT_MAPPING[i]
                image_name = img_file.stem  # 去掉文件扩展名
                
                # 为每个文本框名称查找并替换文本
                for text_box_name in text_box_names:
                    # 按名称查找文本框
                    for shape in slide.shapes:
                        current_shape_name = getattr(shape, 'name', '')
                        if current_shape_name == text_box_name:
                            # 检查是否是文本框
                            if hasattr(shape, 'text_frame') and shape.text_frame:
                                print(f"    📝 正在替换文本框'{text_box_name}'为'{image_name}'")
                                
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
                                else:
                                    # 如果没有段落，直接设置文本（回退方案）
                                    text_frame.text = image_name
                                
                                text_replacements_made += 1
                                print(f"    ✅ 成功更新文本框'{text_box_name}'（已保持格式）")
                                break
                            elif hasattr(shape, 'text'):
                                print(f"    📝 正在替换文本形状'{text_box_name}'为'{image_name}'")
                                shape.text = image_name
                                text_replacements_made += 1
                                print(f"    ✅ 成功更新文本形状'{text_box_name}'")
                                break
        
        # 保存演示文稿
        prs.save(pptx_path)
        print(f"  💾 已保存演示文稿，共替换{replacements_made}个图片，{text_replacements_made}个文本")
        return (replacements_made > 0) or (text_replacements_made > 0)
        
    except Exception as e:
        print(f"  ❌ 更新演示文稿时出错：{e}")
        import traceback
        traceback.print_exc()
        return False

def process_folder(folder_path, template_pptx):
    """处理包含3张图片文件的单个文件夹"""
    print(f"\n📂 正在处理文件夹：{folder_path.name}")
    
    # 检查是否已经存在生成的演示文稿
    existing_pptx = list(folder_path.glob(f"{folder_path.name}_presentation*.pptx"))
    if existing_pptx:
        print(f"  ⏭️  跳过 - 文件夹中已存在生成的演示文稿：{existing_pptx[0].name}")
        return True  # 返回True因为这被认为是"成功处理"（已经处理过）
    
    # 检查必需的输入图片文件
    present_files, missing_files = check_required_files(folder_path)
    
    if missing_files:
        print(f"  ❌ {', '.join(missing_files)}")
        print(f"  📝 每个文件夹必须包含恰好{REQUIRED_IMAGE_COUNT}张图片文件（支持格式：{', '.join(SUPPORTED_EXTENSIONS)}）")
        return False
    
    print(f"  ✅ 找到所需的图片文件：{len(present_files)}/{REQUIRED_IMAGE_COUNT}")
    
    # 显示每个文件将替换的形状和文本
    for i, img_file in enumerate(present_files):
        if i < len(IMAGE_SHAPE_MAPPING):
            shapes = IMAGE_SHAPE_MAPPING[i]
            print(f"  📷 {img_file.name} → {', '.join(shapes)}")
        if i < len(TEXT_REPLACEMENT_MAPPING):
            text_boxes = TEXT_REPLACEMENT_MAPPING[i]
            image_name = img_file.stem
            print(f"  📝 '{image_name}' → {', '.join(text_boxes)}")
    
    # 创建演示文稿副本并替换图片
    print("  🔧 正在更新演示文稿...")
    pptx_copy = copy_template_pptx(template_pptx, folder_path, folder_path.name)
    
    if pptx_copy:
        success = replace_images_in_presentation(pptx_copy, present_files)
        if success:
            print(f"  🎉 成功处理文件夹：{folder_path.name}")
            print(f"  📊 处理的图片总数：{len(present_files)}")
            return True
        else:
            print(f"  ⚠️  在{folder_path.name}中未进行图片或文本替换")
            print("  💡 请确保PowerPoint形状名称和文本框名称匹配预设映射")
            return False
    
    return False

def main():
    """主函数，协调整个处理过程"""
    try:
        # 自动检测工作目录
        working_dir = get_script_directory()
        
        print_separator()
        print("🧾 口袋巾图片处理器")
        print_separator()
        print("此脚本将自动处理图片并更新PowerPoint演示文稿。")
        print(f"工作目录：{working_dir}")
        print()
        print("📋 每个产品文件夹中需要的文件：")
        print(f"  - 恰好{REQUIRED_IMAGE_COUNT}张图片文件（任意名称）")
        print(f"  - 支持格式：{', '.join(SUPPORTED_EXTENSIONS)}")
        print()
        print("🔧 图片形状映射关系（自动分配）：")
        for i, shapes in enumerate(IMAGE_SHAPE_MAPPING, 1):
            print(f"  - 第{i}张图片 → {', '.join(shapes)} ({len(shapes)}个形状)")
        print()
        print("📝 文本替换映射关系（自动分配）：")
        for i, text_boxes in enumerate(TEXT_REPLACEMENT_MAPPING, 1):
            print(f"  - 第{i}张图片名称 → {', '.join(text_boxes)} ({len(text_boxes)}个文本框)")
        print()
        
        # 查找模板PPTX
        template_pptx = find_template_pptx(working_dir)
        if not template_pptx:
            return
        
        print()
        print_separator()
        print("🔧 处理配置")
        print_separator()
        print(f"工作目录：{working_dir}")
        print(f"模板文件：{template_pptx.name}")
        print(f"每个文件夹所需图片文件：{REQUIRED_IMAGE_COUNT}张 (支持格式：{', '.join(SUPPORTED_EXTENSIONS)})")
        
        # 查找要处理的子目录
        subdirs = [d for d in working_dir.iterdir() if d.is_dir() and not d.name.startswith('.')]
        
        if not subdirs:
            print()
            print("❌ 未找到要处理的子目录。")
            print("请为每个产品创建包含图片文件的子文件夹。")
            input("按回车键退出...")
            return
        
        print(f"\n找到{len(subdirs)}个文件夹要处理：")
        for subdir in subdirs:
            print(f"  📁 {subdir.name}")
        
        # 自动开始或确认
        print()
        response = input("🚀 准备开始处理？（按回车继续，输入'n'取消）：").lower()
        if response == 'n':
            print("操作已取消。")
            return
        
        print()
        print_separator()
        print("🚀 开始处理")
        print_separator()
        
        # 处理每个子目录
        successful = 0
        for subdir in subdirs:
            if process_folder(subdir, template_pptx):
                successful += 1
        
        # 最终总结
        print()
        print_separator()
        print("✨ 处理完成")
        print_separator()
        print(f"成功处理：{successful}/{len(subdirs)}个文件夹")
        
        if successful > 0:
            print("\n🎉 您的演示文稿已准备就绪！")
            print("请检查每个文件夹中更新的PowerPoint文件。")
            print("每个文件夹现在包含：")
            print(f"  📁 原始文件：{REQUIRED_IMAGE_COUNT}张图片文件（任意名称，支持格式：{', '.join(SUPPORTED_EXTENSIONS)}）")
            print("  📋 演示文稿：[文件夹名称]_presentation.pptx")
            print("  🔄 已完成图片替换和文本更新")
        else:
            print("\n⚠️  未成功处理任何文件夹。")
            print("请检查您的文件是否正确组织。")
        
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
