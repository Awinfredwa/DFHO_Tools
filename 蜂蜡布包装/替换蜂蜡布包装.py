#!/usr/bin/env python3
"""
蜂蜡布包装图片处理脚本
===================

此脚本自动完成以下任务：
1. 查找文件夹中的单个图片文件（PNG或JPG）
2. 从输入图片生成3个不同尺寸的图片（大、中、小）
3. 在PowerPoint演示文稿中替换指定的图片形状
4. 保持原始图片的尺寸和层级

简易设置：
1. 将此脚本放在您的工作文件夹中（与template.pptx一起）
2. 在子文件夹中放置您的图片文件（PNG或JPG）
3. 每个子文件夹必须包含：一个图片文件，且没有现有的PPTX文件
4. 脚本将生成3个尺寸的图片并替换PowerPoint中的指定形状

专为非技术用户设计 - 只需点击运行！
"""

import os
import sys
import shutil
import tempfile
from PIL import Image
from pptx import Presentation
from pathlib import Path

# 生成的图片尺寸配置
GENERATED_SIZES = {
    "2-1L": (33, 35),  # 大号：33x35厘米
    "2-2M": (25, 28),  # 中号：25x28厘米
    "2-3S": (18, 20),  # 小号：18x20厘米
}

# 图片尺寸到PowerPoint形状名称的映射
IMAGE_SHAPE_MAPPING = {
    "2-1L": ["2-1L"],  # 大号图片的形状
    "2-2M": ["2-2M"],   # 中号图片的形状
    "2-3S": ["2-3S"]              # 小号图片的形状
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

def find_single_image(folder_path):
    """查找文件夹中的单个图片文件（PNG或JPG）"""
    # 支持的图片格式
    image_extensions = ['.png', '.jpg', '.jpeg', '.PNG', '.JPG', '.JPEG']
    
    image_files = []
    for ext in image_extensions:
        image_files.extend(folder_path.glob(f"*{ext}"))
    
    # 检查是否恰好有一个图片文件
    if len(image_files) == 0:
        return None, "未找到图片文件"
    elif len(image_files) > 1:
        return None, f"找到多个图片文件：{', '.join([f.name for f in image_files])}"
    else:
        return image_files[0], None

def has_existing_pptx(folder_path):
    """检查文件夹中是否已有PPTX文件"""
    pptx_files = list(folder_path.glob("*.pptx"))
    return len(pptx_files) > 0, pptx_files

def create_sized_images_from_source(source_image_path, dpi=96):
    """从源图片创建多种尺寸版本（内存中处理，不保存到磁盘）"""
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
            
            # 创建临时文件夹存储处理后的图片
            temp_dir = Path(tempfile.mkdtemp())
            created_files = {}
            
            for size_name, (width_cm, height_cm) in GENERATED_SIZES.items():
                crop_width = cm_to_pixel(width_cm)
                crop_height = cm_to_pixel(height_cm)
                
                # 现在图片应该足够大了，但还是检查一下
                if crop_width > width or crop_height > height:
                    print(f"  ⚠️  意外错误：即使放大后，{size_name}仍然过大，跳过")
                    continue
                
                # 从中心计算裁剪区域
                left = center_x - crop_width // 2
                top = center_y - crop_height // 2
                right = left + crop_width
                bottom = top + crop_height
                
                # 裁剪图片
                cropped_image = image.crop((left, top, right, bottom))
                
                # 保存到临时文件
                temp_file_path = temp_dir / f"{size_name}.png"
                cropped_image.save(temp_file_path)
                created_files[size_name] = temp_file_path
                print(f"  ✅ 已创建{size_name}：{width_cm}x{height_cm}厘米")
            
            return created_files, temp_dir
            
    except Exception as e:
        print(f"  ❌ 从源图片创建尺寸版本时出错：{e}")
        return None, None

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

def replace_images_in_presentation(pptx_path, sized_images):
    """替换PowerPoint演示文稿中的图片"""
    try:
        prs = Presentation(pptx_path)
        replacements_made = 0
        
        # 创建尺寸名称到形状名称列表的映射
        print(f"  🔍 准备替换以下图片映射：")
        for size_name, img_path in sized_images.items():
            if size_name in IMAGE_SHAPE_MAPPING:
                shape_names = IMAGE_SHAPE_MAPPING[size_name]
                print(f"    {size_name} → {', '.join(shape_names)}")
        
        for slide_idx, slide in enumerate(prs.slides):
            print(f"  📄 正在处理幻灯片 {slide_idx + 1}")
            
            for size_name, img_path in sized_images.items():
                if size_name not in IMAGE_SHAPE_MAPPING:
                    continue
                    
                shape_names = IMAGE_SHAPE_MAPPING[size_name]
                
                # 检查图片是否存在
                if not img_path.exists():
                    print(f"    ❌ 未找到图片：{img_path}")
                    continue
                
                # 为每个形状名称查找并替换
                for shape_name in shape_names:
                    # 按名称查找形状
                    for shape in slide.shapes:
                        if hasattr(shape, 'name') and shape.name == shape_name and shape.shape_type == 13:  # 13表示图片类型
                            print(f"    🔄 正在替换形状'{shape_name}'为{size_name}")
                            
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
                            new_pic = slide.shapes.add_picture(str(img_path), left, top, width, height)
                            
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
                            
                            print(f"    ✅ 成功替换'{shape_name}'（已保持正确的XML结构和层级）")
                            replacements_made += 1
                            break
        
        # 保存演示文稿
        prs.save(pptx_path)
        print(f"  💾 已保存演示文稿，共替换{replacements_made}个图片")
        return replacements_made > 0
        
    except Exception as e:
        print(f"  ❌ 更新演示文稿时出错：{e}")
        return False

def cleanup_temp_files(temp_dir):
    """清理临时文件"""
    try:
        if temp_dir and temp_dir.exists():
            shutil.rmtree(temp_dir)
            print(f"  🧹 已清理临时文件")
    except Exception as e:
        print(f"  ⚠️  清理临时文件时出现警告：{e}")

def process_folder(folder_path, template_pptx):
    """处理包含单个图片文件的文件夹"""
    print(f"\n📂 正在处理文件夹：{folder_path.name}")
    
    # 检查是否已经存在PPTX文件（跳过已处理的文件夹）
    has_pptx, existing_pptx = has_existing_pptx(folder_path)
    if has_pptx:
        print(f"  ⏭️  跳过 - 文件夹中已存在PPTX文件：{', '.join([f.name for f in existing_pptx])}")
        return True  # 返回True因为这被认为是"成功处理"（已经处理过）
    
    # 查找单个图片文件
    image_file, error_msg = find_single_image(folder_path)
    
    if image_file is None:
        print(f"  ❌ {error_msg}")
        print(f"  📝 每个文件夹必须包含恰好一个图片文件（PNG或JPG格式）")
        return False
    
    print(f"  ✅ 找到图片文件：{image_file.name}")
    
    # 从源图片创建3个尺寸的版本
    print("  🔧 正在生成不同尺寸的图片...")
    sized_images, temp_dir = create_sized_images_from_source(image_file)
    
    if sized_images is None:
        print(f"  ❌ 无法生成尺寸图片")
        return False
    
    try:
        # 显示每个尺寸将替换的形状
        for size_name, img_path in sized_images.items():
            if size_name in IMAGE_SHAPE_MAPPING:
                shapes = IMAGE_SHAPE_MAPPING[size_name]
                print(f"  📷 {size_name} → {', '.join(shapes)}")
        
        # 创建演示文稿副本并替换图片
        print("  🔧 正在更新演示文稿...")
        pptx_copy = copy_template_pptx(template_pptx, folder_path, folder_path.name)
        
        if pptx_copy:
            success = replace_images_in_presentation(pptx_copy, sized_images)
            if success:
                print(f"  🎉 成功处理文件夹：{folder_path.name}")
                print(f"  📊 生成的尺寸数量：{len(sized_images)}")
                return True
            else:
                print(f"  ⚠️  在{folder_path.name}中未进行图片替换")
                print("  💡 请确保PowerPoint形状名称匹配预设映射")
                return False
        
        return False
    
    finally:
        # 清理临时文件
        cleanup_temp_files(temp_dir)

def main():
    """主函数，协调整个处理过程"""
    try:
        # 自动检测工作目录
        working_dir = get_script_directory()
        
        print_separator()
        print("🧾 蜂蜡布包装图片处理器")
        print_separator()
        print("此脚本将自动处理图片并更新PowerPoint演示文稿。")
        print(f"工作目录：{working_dir}")
        print()
        print("📋 每个产品文件夹中的要求：")
        print("  - 恰好一个图片文件（PNG或JPG格式）")
        print("  - 没有现有的PPTX文件")
        print()
        print("🔧 生成的图片尺寸：")
        for size_name, (w, h) in GENERATED_SIZES.items():
            print(f"  - {size_name}：{w}x{h}厘米")
        print()
        print("🎯 图片形状映射关系：")
        for size_name, shapes in IMAGE_SHAPE_MAPPING.items():
            print(f"  - {size_name} → {', '.join(shapes)}")
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
        print(f"支持的图片格式：PNG, JPG")
        
        # 查找要处理的子目录
        subdirs = [d for d in working_dir.iterdir() if d.is_dir() and not d.name.startswith('.')]
        
        if not subdirs:
            print()
            print("❌ 未找到要处理的子目录。")
            print("请为每个产品创建包含单个图片文件的子文件夹。")
            input("按回车键退出...")
            return
        
        # 过滤出符合条件的文件夹（有一个图片且没有pptx）
        valid_folders = []
        for subdir in subdirs:
            image_file, error_msg = find_single_image(subdir)
            has_pptx, _ = has_existing_pptx(subdir)
            
            if image_file is not None and not has_pptx:
                valid_folders.append(subdir)
        
        print(f"\n找到{len(valid_folders)}个符合条件的文件夹要处理：")
        for subdir in valid_folders:
            image_file, _ = find_single_image(subdir)
            print(f"  📁 {subdir.name} (图片：{image_file.name})")
        
        if not valid_folders:
            print("❌ 没有找到符合条件的文件夹。")
            print("每个文件夹必须包含恰好一个图片文件且没有现有的PPTX文件。")
            input("按回车键退出...")
            return
        
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
        
        # 处理每个符合条件的子目录
        successful = 0
        for subdir in valid_folders:
            if process_folder(subdir, template_pptx):
                successful += 1
        
        # 最终总结
        print()
        print_separator()
        print("✨ 处理完成")
        print_separator()
        print(f"成功处理：{successful}/{len(valid_folders)}个文件夹")
        
        if successful > 0:
            print("\n🎉 您的演示文稿已准备就绪！")
            print("请检查每个文件夹中更新的PowerPoint文件。")
            print("每个文件夹现在包含：")
            print("  📁 原始图片文件（PNG或JPG）")
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
