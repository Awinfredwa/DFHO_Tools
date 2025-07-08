#!/usr/bin/env python3
"""
蜂蜡布图片处理脚本
================

此脚本自动完成以下任务：
1. 剪裁图片去除空白区域
2. 从源图片创建多种尺寸版本
3. 在PowerPoint演示文稿中替换图片

简易设置：
1. 将此脚本放在您的工作文件夹中（与template.pptx一起）
2. 在子文件夹中整理您的PNG文件
3. 每个子文件夹必须包含：1-1L.png, 1-2M.png, 1-3S.png, 和 2.png
4. 脚本将从2.png创建2-1L.png, 2-2M.png, 2-3S.png
5. 双击"Run_Image_Processor.bat"开始！

专为非技术用户设计 - 只需点击运行！
"""

import os
import sys
import shutil
from PIL import Image
from pptx import Presentation
from pathlib import Path

# 用户需要提供的PNG文件
REQUIRED_INPUT_FILES = ["1-1L.png", "1-2M.png", "1-3S.png", "2.png"]

# 从2.png生成的文件及其尺寸（单位：厘米）
GENERATED_SIZES = {
    "2-1L.png": (33, 35),  # 大号：33x35厘米
    "2-2M.png": (25, 28),  # 中号：25x28厘米
    "2-3S.png": (18, 20),  # 小号：18x20厘米
}

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

def crop_image_whitespace(image_path):
    """移除图片周围的空白区域"""
    try:
        with Image.open(image_path) as img:
            bbox = img.getbbox()
            if bbox:
                cropped_img = img.crop(bbox)
                cropped_img.save(image_path)
                print(f"  ✂️  已移除空白区域：{image_path.name}")
                return True
            else:
                print(f"  ⚠️  未发现空白区域：{image_path.name}")
                return True
    except Exception as e:
        print(f"  ❌ 处理图片{image_path.name}时出错：{e}")
        return False

def check_required_files(folder_path):
    """检查文件夹中是否存在所有必需的输入PNG文件"""
    missing_files = []
    present_files = []
    
    for required_file in REQUIRED_INPUT_FILES:
        file_path = folder_path / required_file
        if file_path.exists():
            present_files.append(file_path)
        else:
            missing_files.append(required_file)
    
    return present_files, missing_files

def create_sized_images_from_source(source_image_path, dpi=96):
    """从源图片（2.png）创建多种尺寸版本"""
    cm_to_pixel = lambda cm: int(cm * dpi / 2.54)
    
    try:
        with Image.open(source_image_path) as original_image:
            print(f"  📐 源图片尺寸：{original_image.size[0]}x{original_image.size[1]}像素")
            
            # 计算所有需要的最大尺寸
            required_widths = [cm_to_pixel(w) for _, (w, h) in GENERATED_SIZES.items()]
            required_heights = [cm_to_pixel(h) for _, (w, h) in GENERATED_SIZES.items()]
            min_width_needed = max(required_widths)
            min_height_needed = max(required_heights)
            
            current_width, current_height = original_image.size
            
            # 检查是否需要放大图片
            if current_width < min_width_needed or current_height < min_height_needed:
                # 计算需要的缩放比例（保持宽高比）
                scale_x = min_width_needed / current_width
                scale_y = min_height_needed / current_height
                scale_factor = max(scale_x, scale_y)
                
                new_width = int(current_width * scale_factor)
                new_height = int(current_height * scale_factor)
                
                print(f"  🔍 图片太小，正在放大...")
                print(f"  📏 需要最小尺寸：{min_width_needed}x{min_height_needed}像素")
                print(f"  ⬆️  放大到：{new_width}x{new_height}像素（缩放{scale_factor:.2f}倍）")
                
                # 使用高质量重采样放大图片
                image = original_image.resize((new_width, new_height), Image.Resampling.LANCZOS)
            else:
                print(f"  ✅ 图片尺寸足够大")
                image = original_image.copy()
            
            width, height = image.size
            center_x, center_y = width // 2, height // 2
            
            created_files = []
            
            for output_name, (width_cm, height_cm) in GENERATED_SIZES.items():
                crop_width = cm_to_pixel(width_cm)
                crop_height = cm_to_pixel(height_cm)
                
                # 现在图片应该足够大了，但还是检查一下
                if crop_width > width or crop_height > height:
                    print(f"  ⚠️  意外错误：即使放大后，{output_name}仍然过大，跳过")
                    continue
                
                # 从中心计算裁剪区域
                left = center_x - crop_width // 2
                top = center_y - crop_height // 2
                right = left + crop_width
                bottom = top + crop_height
                
                # 裁剪并保存
                cropped_image = image.crop((left, top, right, bottom))
                output_path = source_image_path.parent / output_name
                cropped_image.save(output_path)
                created_files.append(output_path)
                print(f"  ✅ 已创建{output_name}：{width_cm}x{height_cm}厘米")
            
            return created_files
            
    except Exception as e:
        print(f"  ❌ 从源图片创建尺寸版本时出错：{e}")
        return None

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
    """替换PowerPoint演示文稿中的图片"""
    try:
        prs = Presentation(pptx_path)
        replacements_made = 0
        
        # 创建形状名称到图片文件的映射
        image_map = {}
        for img_file in image_files:
            shape_name = img_file.stem  # 不带扩展名的文件名
            image_map[shape_name] = img_file
        
        print(f"  🔍 查找要替换的形状：{list(image_map.keys())}")
        
        for slide_idx, slide in enumerate(prs.slides):
            print(f"  📄 正在处理幻灯片 {slide_idx + 1}")
            
            for shape_name, new_image_path in image_map.items():
                # 检查新图片是否存在
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
                        shape_index = slide.shapes.index(shape)  # 获取形状的索引
                        
                        # 移除原始形状
                        sp = shape._element  # 获取形状的XML元素
                        sp.getparent().remove(sp)  # 从幻灯片中移除形状
                        
                        # 添加新图片并保持原始位置和大小
                        new_pic = slide.shapes.add_picture(str(new_image_path), left, top, width, height)
                        
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
                        
                        print(f"    ✅ 成功替换'{shape_name}'（已保持正确的XML结构）")
                        replacements_made += 1
                        break
                else:
                    # 如果for循环完成而没有break（未找到形状）
                    continue  # 在此幻灯片上未找到形状，没关系
        
        # 保存演示文稿
        prs.save(pptx_path)
        print(f"  💾 已保存演示文稿，共替换{replacements_made}个图片")
        return replacements_made > 0
        
    except Exception as e:
        print(f"  ❌ 更新演示文稿时出错：{e}")
        return False

def process_folder(folder_path, template_pptx):
    """处理包含4个必需输入PNG图片的单个文件夹"""
    print(f"\n📂 正在处理文件夹：{folder_path.name}")
    
    # === 更可靠的演示文稿检测方法（适用于网络共享） ===
    # 使用手动文件列表而不是glob，以避免网络共享问题
    existing_pptx = []
    try:
        # 获取文件夹中所有文件的列表
        for file_path in folder_path.iterdir():
            if file_path.is_file() and file_path.suffix.lower() == '.pptx':
                # 检查是否是生成的演示文稿文件
                if file_path.stem.startswith(f"{folder_path.name}_presentation"):
                    existing_pptx.append(file_path)
    except Exception as e:
        print(f"  ⚠️  访问文件夹时出错：{e}")
        # 继续处理，不要因为文件访问问题而停止
    
    if existing_pptx:
        print(f"  ⏭️  跳过 - 文件夹中已存在生成的演示文稿：{existing_pptx[0].name}")
        return True  # 返回True因为这被认为是"成功处理"（已经处理过）
    
    # 检查必需的输入PNG文件
    present_files, missing_files = check_required_files(folder_path)
    
    if missing_files:
        print(f"  ❌ 缺少必需的输入文件：{', '.join(missing_files)}")
        print(f"  📝 每个文件夹必须包含：{', '.join(REQUIRED_INPUT_FILES)}")
        return False
    
    print(f"  ✅ 找到所有必需的输入文件：{len(present_files)}/4")
    
    # 步骤1：移除所有输入PNG文件的空白区域
    print("  🔧 步骤1：移除输入文件的空白区域...")
    for png_file in present_files:
        crop_image_whitespace(png_file)
    
    # 步骤2：从2.png创建尺寸版本
    source_image = folder_path / "2.png"
    print("  🔧 步骤2：从2.png创建尺寸版本...")
    created_files = create_sized_images_from_source(source_image)
    
    if created_files is None:
        return False
    
    # 合并所有图片文件（输入+生成）
    all_image_files = present_files + created_files
    
    # 步骤3：创建演示文稿副本并替换图片
    print("  🔧 步骤3：更新演示文稿...")
    pptx_copy = copy_template_pptx(template_pptx, folder_path, folder_path.name)
    
    if pptx_copy:
        success = replace_images_in_presentation(pptx_copy, all_image_files)
        if success:
            print(f"  🎉 成功处理文件夹：{folder_path.name}")
            print(f"  📊 处理的图片总数：{len(all_image_files)}（4个输入 + {len(created_files)}个生成）")
            return True
        else:
            print(f"  ⚠️  在{folder_path.name}中未进行图片替换")
            print("  💡 请确保PowerPoint形状名称匹配：")
            print("     输入文件：1-1L, 1-2M, 1-3S, 2")
            print("     生成文件：2-1L, 2-2M, 2-3S")
            return False
    
    return False

def main():
    """主函数，协调整个处理过程"""
    try:
        # 自动检测工作目录
        working_dir = get_script_directory()
        
        print_separator()
        print("🐝 蜂蜡布图片处理器")
        print_separator()
        print("此脚本将自动处理图片并更新PowerPoint演示文稿。")
        print(f"工作目录：{working_dir}")
        print()
        print("📋 每个产品文件夹中必需的输入文件：")
        for file in REQUIRED_INPUT_FILES:
            print(f"  - {file}")
        print()
        print("🔧 生成的文件（从2.png创建）：")
        for file, (w, h) in GENERATED_SIZES.items():
            print(f"  - {file}（{w}x{h}厘米）")
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
        print(f"每个文件夹所需输入文件：{len(REQUIRED_INPUT_FILES)}")
        print(f"每个文件夹生成文件：{len(GENERATED_SIZES)}")
        
        # 查找要处理的子目录
        subdirs = [d for d in working_dir.iterdir() if d.is_dir()]
        
        if not subdirs:
            print()
            print("❌ 未找到要处理的子目录。")
            print("请为每个产品创建包含必需PNG文件的子文件夹。")
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
            print("  📁 原始文件：1-1L.png, 1-2M.png, 1-3S.png, 2.png")
            print("  ⚡ 生成文件：2-1L.png, 2-2M.png, 2-3S.png")
            print("  📋 演示文稿：[文件夹名称]_presentation.pptx")
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