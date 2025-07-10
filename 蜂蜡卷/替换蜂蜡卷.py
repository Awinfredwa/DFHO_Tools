
import os
import sys
import shutil
from PIL import Image
from pptx import Presentation
from pathlib import Path

# 支持的图片文件扩展名
SUPPORTED_EXTENSIONS = ['.png', '.jpg', '.jpeg']

# 用户需要提供的图片文件（不带扩展名）
REQUIRED_INPUT_FILES = ["1", "2", "3"]

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

def find_image_file(folder_path, base_name):
    """查找具有指定基本名称的图片文件（不区分大小写，支持多种扩展名）"""
    for ext in SUPPORTED_EXTENSIONS:
        # 尝试精确匹配
        file_path = folder_path / f"{base_name}{ext}"
        if file_path.exists():
            return file_path
        
        # 尝试不区分大小写匹配
        for f in folder_path.glob(f"*{ext}"):
            if f.stem.lower() == base_name.lower():
                return f
    return None

def check_required_files(folder_path):
    """检查文件夹中是否存在所有必需的输入图片文件"""
    missing_files = []
    present_files = []
    
    for base_name in REQUIRED_INPUT_FILES:
        file_path = find_image_file(folder_path, base_name)
        if file_path:
            present_files.append(file_path)
        else:
            missing_files.append(base_name)
    
    return present_files, missing_files

def find_additional_image(folder_path, required_files):
    """查找文件夹中除必需文件外的第一个图片文件"""
    required_paths = {str(f).lower() for f in required_files}
    
    for ext in SUPPORTED_EXTENSIONS:
        for img_file in folder_path.glob(f"*{ext}"):
            if str(img_file).lower() not in required_paths and not img_file.stem.endswith('_temp'):
                return img_file
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

def replace_images_in_presentation(pptx_path, image_files, folder_name=None):
    """替换PowerPoint演示文稿中的图片和文本"""
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
                image_map['1'] = img_file
            elif base_name == '2':
                image_map['2'] = img_file
            elif base_name == '3':
                image_map['3'] = img_file
            else:
                # 保存额外图片，稍后用于替换"图片 14"
                additional_image = img_file
        
        # 如果有额外图片，用于替换"图片 14"
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
        
        # 处理文本替换（使用额外图片的文件名）
        print(f"\n🔍 文本替换调试信息:")
        
        # 查找额外图片（文件名不是1,2,3的）
        additional_image = None
        for img_file in image_files:
            base_name = img_file.stem
            if base_name not in ['1', '2', '3']:
                additional_image = img_file
                break
        
        if additional_image:
            # 获取额外图片的文件名（不带扩展名）
            text_to_use = additional_image.stem
            
            print(f"   将使用文本: '{text_to_use}' 进行替换")
            print(f"   正在查找文本框: 'name'")
                
            # 定义要替换的文本框名称列表
            text_boxes_to_replace = ['name']
            text_box_found = False
            
            for slide_idx, slide in enumerate(prs.slides):
                print(f"  📄 正在检查幻灯片 {slide_idx + 1} 中的文本框...")
                
                # 查找并替换所有指定的文本框
                for text_box_name in text_boxes_to_replace:
                    for shape in slide.shapes:
                        shape_name = getattr(shape, 'name', '无名称')
                        print(f"    检查形状: '{shape_name}' (类型: {type(shape).__name__})")
                        
                        if hasattr(shape, 'name') and shape.name == text_box_name:
                            print(f"    ✅ 找到匹配的文本框: '{shape.name}'")
                            text_box_found = True
                            
                            # 检查是否是文本框
                            if hasattr(shape, 'text_frame') and shape.text_frame:
                                print(f"    📝 正在替换文本框'{shape.name}'为'{text_to_use}'")
                            
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
                                new_run.text = text_to_use
                                
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
                                print(f"    ✅ 成功更新文本框'{text_box_name}'（已保持格式）")
                                text_replacements_made += 1
                                break
                            elif hasattr(shape, 'text'):
                                print(f"    📝 正在替换文本形状'{text_box_name}'为'{text_to_use}'")
                                shape.text = folder_name
                                text_replacements_made += 1
                                print(f"    ✅ 成功更新文本形状'{text_box_name}'")
                                break
        
        # 保存演示文稿
        prs.save(pptx_path)
        print(f"\n📊 替换统计:")
        print(f"   - 图片替换: {replacements_made} 个")
        print(f"   - 文本替换: {text_replacements_made} 个")
        print(f"   - 额外图片: {'已使用' if additional_image else '未找到'}")
        if additional_image:
            print(f"   - 使用的额外图片: {additional_image.name}")
            print(f"   - 文本替换内容: '{text_to_use}'")
        return (replacements_made > 0) or (text_replacements_made > 0)
        
    except Exception as e:
        print(f"  ❌ 更新演示文稿时出错：{e}")
        return False

def process_folder(folder_path, template_pptx):
    """处理包含必需输入图片文件的单个文件夹"""
    print(f"\n📂 正在处理文件夹：{folder_path.name}")
    
    # 检查是否已存在演示文稿
    pptx_files = list(folder_path.glob("*.pptx"))
    if pptx_files:
        print(f"  ⚠️  文件夹中已存在PPTX文件，跳过处理")
        return False
    
    # 检查必需的文件
    present_files, missing_files = check_required_files(folder_path)
    
    if missing_files:
        print(f"  ❌ 缺少以下必需文件：{', '.join(missing_files)}")
        return False
    
    print(f"  ✅ 找到所有必需的文件")
    
    # 查找额外图片（如果有）
    additional_image = find_additional_image(folder_path, present_files)
    if additional_image:
        print(f"  ➕ 找到额外图片：{additional_image.name}")
    
    # 创建PPTX的临时副本
    print("  📝 正在准备演示文稿...")
    pptx_copy_path = copy_template_pptx(template_pptx, folder_path, folder_path.name)
    
    if not pptx_copy_path:
        return False
    
    # 使用原始文件
    image_files = present_files.copy()
    
    # 如果需要，添加额外图片到处理列表
    if additional_image:
        image_files.append(additional_image)
    
    # 替换演示文稿中的图片和文本
    print("  🖼️  正在更新演示文稿...")
    success = replace_images_in_presentation(pptx_copy_path, image_files, folder_path.name)
    
    # 如果处理失败，删除创建的PPTX文件
    if not success and pptx_copy_path.exists():
        try:
            pptx_copy_path.unlink()
            print(f"  🗑️  已删除未完成的演示文稿")
        except Exception as e:
            print(f"  ⚠️  删除未完成的演示文稿时出错：{e}")
    
    if success:
        print(f"  🎉 成功处理文件夹：{folder_path.name}")
        print(f"  📊 处理的图片总数：{len(present_files)}")
    else:
        print(f"  ⚠️  在 {folder_path.name} 中未进行图片替换")
        print("  💡 请确保PowerPoint形状名称与图片文件名匹配")
    
    return success

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
            print(f"  - {file} (支持 {', '.join(ext.upper() for ext in SUPPORTED_EXTENSIONS)})")
        print("  - 第4个文件（可选，任意名称）")
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
        print(f"支持的图片格式：{', '.join(ext.upper() for ext in SUPPORTED_EXTENSIONS)}")
        
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
            print("  📁 原始文件：1, 2, 3 (支持PNG/JPG)，以及第4个文件（可选）")
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