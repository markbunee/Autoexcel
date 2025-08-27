"""
from autoexcel.file_classify import classify_files

# 分类所有文件类型
classify_files("/path/to/source", "/path/to/target", None)

# 只分类特定文件类型
classify_files("/path/to/source", "/path/to/target", ["word", "excel", "pdf"])

将一个跟目录文件夹下的每一个文件划分到指定文件夹下
"""

import os
import shutil
import argparse
from pathlib import Path
import re
from collections import defaultdict

# 定义文件类型和对应文件夹的映射关系
FILE_TYPES = {
    'word': ['doc', 'docx', 'dot', 'dotx', 'docm', 'dotm'],
    'excel': ['xls', 'xlsx', 'xlt', 'xltx', 'xlsm', 'xltm'],
    'pdf': ['pdf'],
    'text': ['txt'],
    'image': ['png', 'jpg', 'jpeg', 'gif', 'bmp', 'tiff', 'svg'],
    'archive': ['zip', 'rar', '7z', 'tar', 'gz'],
    'powerpoint': ['ppt', 'pptx', 'pot', 'potx', 'ppsx', 'pptm', 'potm', 'ppsm'],
    'code': ['py', 'js', 'html', 'css', 'java', 'cpp', 'c', 'h', 'php', 'sql'],
    'audio': ['mp3', 'wav', 'flac', 'aac', 'ogg', 'wma'],
    'video': ['mp4', 'avi', 'mkv', 'mov', 'wmv', 'flv', 'webm'],
    'med_imge': ['nii.gz'],
    'other': ['json']
}

# 细分类型定义
SUB_TYPES = {
    'image': {
        'png_image': ['png'],
        'jpg_image': ['jpg', 'jpeg'],
        'gif_image': ['gif'],
        'bmp_image': ['bmp'],
        'tiff_image': ['tiff'],
        'svg_image': ['svg']
    },
    'med_imge': {
        'nii_gz': ['nii.gz']
    }
}

# 文件名模式定义
NAME_PATTERNS = {
    'date': r'\d{4}-\d{2}-\d{2}',  # 匹配日期格式(YYYY-MM-DD)
    'number': r'\d+',             # 匹配数字
    'report': r'(?i)report'       # 匹配"report"，不区分大小写
}

def create_folders(base_path, selected_types=None, sub_types=None, name_patterns=None, keywords=None):
    """创建所有需要的文件夹"""
    # 创建主类型文件夹
    types_to_create = FILE_TYPES.keys() if selected_types is None else selected_types
    for folder_name in types_to_create:
        if folder_name in FILE_TYPES:  # 确保是有效的文件类型
            folder_path = os.path.join(base_path, folder_name)
            os.makedirs(folder_path, exist_ok=True)
    
    # 创建子类型文件夹
    if sub_types:
        for main_type, sub_type_dict in SUB_TYPES.items():
            if selected_types is None or main_type in selected_types:
                for sub_type_name in sub_type_dict.keys():
                    if sub_type_name in sub_types:
                        folder_path = os.path.join(base_path, main_type, sub_type_name)
                        os.makedirs(folder_path, exist_ok=True)
    
    # 创建文件名模式文件夹
    if name_patterns:
        for pattern_name in name_patterns.keys():
            folder_path = os.path.join(base_path, 'name_pattern', pattern_name)
            os.makedirs(folder_path, exist_ok=True)
    
    # 创建关键词文件夹
    if keywords:
        for keyword in keywords:
            folder_path = os.path.join(base_path, 'keyword', keyword)
            os.makedirs(folder_path, exist_ok=True)

def get_file_type(extension):
    """根据文件扩展名获取文件类型"""
    extension = extension.lower().lstrip('.')
    
    for file_type, extensions in FILE_TYPES.items():
        if extension in extensions:
            return file_type
    
    return 'other'  # 未识别的文件类型

def get_sub_file_type(extension, main_type):
    """根据文件扩展名获取子文件类型"""
    extension = extension.lower().lstrip('.')
    
    if main_type in SUB_TYPES:
        for sub_type_name, extensions in SUB_TYPES[main_type].items():
            if extension in extensions:
                return sub_type_name
    
    return None

def move_file(file_path, destination_folder, output_callback=None):
    """移动文件到目标文件夹"""
    try:
        # 确保目标文件夹存在
        os.makedirs(destination_folder, exist_ok=True)
        
        # 构造目标文件路径
        destination_path = os.path.join(destination_folder, file_path.name)
        
        # 如果目标文件已存在，添加序号
        counter = 1
        while os.path.exists(destination_path):
            name, ext = os.path.splitext(file_path.name)
            destination_path = os.path.join(destination_folder, f"{name}_{counter}{ext}")
            counter += 1
            # 防止无限循环
            if counter > 1000:
                error_msg = f"无法移动文件 {file_path}: 目标文件名冲突过多"
                if output_callback:
                    output_callback(error_msg)
                else:
                    print(error_msg)
                return False, None
        
        shutil.copy(str(file_path), destination_path)
        success_msg = f"已移动: {file_path} -> {destination_path}"
        if output_callback:
            output_callback(success_msg)
        else:
            print(success_msg)
        return True, destination_path
    except Exception as e:
        error_msg = f"移动文件失败 {file_path}: {str(e)}"
        if output_callback:
            output_callback(error_msg)
        else:
            print(error_msg)
        return False, None

def classify_files(source_path, target_path, selected_types=None, name_patterns=None, keywords=None, output_callback=None):
    """
    分类指定目录下的文件
    
    Args:
        source_path (str): 源文件目录路径
        target_path (str): 目标目录路径
        selected_types (list): 要分类的文件类型列表，None表示全部分类
        name_patterns (dict): 文件名模式字典，格式为 {'pattern_name': 'regex_pattern'}
        keywords (list): 关键词列表
        output_callback (function): 输出回调函数，用于将日志信息传递给GUI
    """
    # 检查源目录是否存在
    if not os.path.exists(source_path):
        error_msg = f"源目录 {source_path} 不存在"
        if output_callback:
            output_callback(error_msg)
        else:
            print(error_msg)
        raise FileNotFoundError(error_msg)
    
    # 提取主类型和子类型
    main_types = []
    sub_types = []
    
    if selected_types:
        for file_type in selected_types:
            if '.' in file_type:
                main_type, sub_type = file_type.split('.', 1)
                main_types.append(main_type)
                sub_types.append(file_type)  # 保持完整名称
            else:
                main_types.append(file_type)
    
    # 创建文件夹
    create_folders(target_path, main_types if main_types else selected_types, sub_types, name_patterns, keywords)
    
    # 预先收集所有文件，提高处理效率
    source = Path(source_path)
    all_files = [f for f in source.iterdir() if f.is_file()]
    
    # 按分类类型组织文件，减少重复操作
    name_pattern_files = defaultdict(list)
    keyword_files = defaultdict(list)
    type_files = defaultdict(list)
    
    # 第一遍扫描：分类文件
    for file_path in all_files:
        filename = file_path.name
        extension = file_path.suffix
        
        # 按文件名模式分类
        if name_patterns:
            for pattern_name, pattern in name_patterns.items():
                if re.search(pattern, filename):
                    name_pattern_files[pattern_name].append(file_path)
        
        # 按关键词分类
        if keywords:
            for keyword in keywords:
                if keyword in filename:
                    keyword_files[keyword].append(file_path)
        
        # 按文件类型分类
        file_type = get_file_type(extension)
        type_files[file_type].append(file_path)
    
    # 处理按文件名模式分类的文件（优先级最高）
    processed_files = set()
    destination_folders = set()
    moved_files_count = 0
    if name_patterns:
        for pattern_name, files in name_pattern_files.items():
            destination_folder = os.path.join(target_path, 'name_pattern', pattern_name)
            destination_folders.add(destination_folder)
            for file_path in files:
                success, dest_path = move_file(file_path, destination_folder, output_callback)
                if success:
                    moved_files_count += 1
                processed_files.add(str(file_path))
    
    # 处理按关键词分类的文件（优先级次之）
    if keywords:
        for keyword, files in keyword_files.items():
            destination_folder = os.path.join(target_path, 'keyword', keyword)
            destination_folders.add(destination_folder)
            for file_path in files:
                # 只处理尚未被处理的文件
                if str(file_path) not in processed_files:
                    success, dest_path = move_file(file_path, destination_folder, output_callback)
                    if success:
                        moved_files_count += 1
                    processed_files.add(str(file_path))
    
    # 处理按文件类型分类的文件（优先级最低）
    other_files = []
    for file_type, files in type_files.items():
        for file_path in files:
            # 只处理尚未被处理的文件
            if str(file_path) not in processed_files:
                # 如果指定了文件类型，则只处理这些类型的文件
                if selected_types is not None:
                    # 检查是否匹配主类型
                    main_type_match = file_type in main_types if main_types else True
                    
                    # 检查是否匹配子类型
                    sub_file_type = None
                    if file_type in SUB_TYPES:
                        sub_file_type = get_sub_file_type(file_path.suffix, file_type)
                    
                    sub_type_match = False
                    if sub_file_type:
                        full_sub_type = f"{file_type}.{sub_file_type}"
                        sub_type_match = full_sub_type in sub_types if sub_types else True
                    
                    # 如果既不匹配主类型也不匹配子类型，则跳过
                    if not main_type_match and not sub_type_match:
                        other_files.append(file_path)
                        continue
                
                # 确定目标文件夹
                sub_file_type = None
                if file_type in SUB_TYPES:
                    sub_file_type = get_sub_file_type(file_path.suffix, file_type)
                
                if selected_types and sub_file_type:
                    full_sub_type = f"{file_type}.{sub_file_type}"
                    if full_sub_type in sub_types:
                        destination_folder = os.path.join(target_path, file_type, sub_file_type)
                    else:
                        destination_folder = os.path.join(target_path, file_type)
                else:
                    destination_folder = os.path.join(target_path, file_type)
                
                destination_folders.add(destination_folder)
                success, dest_path = move_file(file_path, destination_folder, output_callback)
                if success:
                    moved_files_count += 1
                processed_files.add(str(file_path))
    
    # 处理未分类的文件（放入"其他"文件夹）
    for file_path in other_files:
        if str(file_path) not in processed_files:
            other_folder = os.path.join(target_path, 'other')
            destination_folders.add(other_folder)
            os.makedirs(other_folder, exist_ok=True)
            success, dest_path = move_file(file_path, other_folder, output_callback)
            if success:
                moved_files_count += 1
    
    # 输出操作总结
    summary_msg = f"操作成功完成！总共处理了 {len(all_files)} 个文件，实际移动了 {moved_files_count} 个文件"
    folder_msg = "文件已移动到以下文件夹:"
    if output_callback:
        output_callback(summary_msg)
        output_callback(folder_msg)
        for folder in sorted(destination_folders):
            output_callback(f"  {folder}")
    else:
        print(summary_msg)
        print(folder_msg)
        for folder in sorted(destination_folders):
            print(f"  {folder}")
    
    return True

def classify_files_by_keywords(source_path, target_path, keywords, output_callback=None):
    """
    根据关键词分类文件（独立功能）
    
    Args:
        source_path (str): 源文件目录路径
        target_path (str): 目标目录路径
        keywords (list): 关键词列表
        output_callback (function): 输出回调函数，用于将日志信息传递给GUI
    """
    # 检查源目录是否存在
    if not os.path.exists(source_path):
        error_msg = f"源目录 {source_path} 不存在"
        if output_callback:
            output_callback(error_msg)
        else:
            print(error_msg)
        raise FileNotFoundError(error_msg)
    
    # 只创建关键词文件夹和other文件夹
    for keyword in keywords:
        folder_path = os.path.join(target_path, 'keyword', keyword)
        os.makedirs(folder_path, exist_ok=True)
    
    # 创建other文件夹用于存放未匹配的文件
    other_folder = os.path.join(target_path, 'other')
    os.makedirs(other_folder, exist_ok=True)
    
    # 预先收集所有文件
    source = Path(source_path)
    all_files = [f for f in source.iterdir() if f.is_file()]
    
    # 按关键词组织文件
    keyword_files = defaultdict(list)
    other_files = []  # 未匹配的文件
    
    for file_path in all_files:
        filename = file_path.name
        matched = False
        for keyword in keywords:
            if keyword in filename:
                keyword_files[keyword].append(file_path)
                matched = True
                break  # 一个文件只匹配一个关键词
        
        if not matched:
            other_files.append(file_path)
    
    # 处理按关键词分类的文件
    destination_folders = set()
    moved_files_count = 0
    
    # 移动匹配关键词的文件
    for keyword, files in keyword_files.items():
        destination_folder = os.path.join(target_path, 'keyword', keyword)
        destination_folders.add(destination_folder)
        for file_path in files:
            success, dest_path = move_file(file_path, destination_folder, output_callback)
            if success:
                moved_files_count += 1
    
    # 移动未匹配的文件到"其他"文件夹
    if other_files:
        destination_folders.add(other_folder)
        for file_path in other_files:
            success, dest_path = move_file(file_path, other_folder, output_callback)
            if success:
                moved_files_count += 1
    
    # 输出操作总结
    summary_msg = f"关键词分类操作成功完成！总共处理了 {len(all_files)} 个文件，实际移动了 {moved_files_count} 个文件"
    folder_msg = "文件已移动到以下文件夹:"
    if output_callback:
        output_callback(summary_msg)
        output_callback(folder_msg)
        for folder in sorted(destination_folders):
            output_callback(f"  {folder}")
    else:
        print(summary_msg)
        print(folder_msg)
        for folder in sorted(destination_folders):
            print(f"  {folder}")
    
    return True

def classify_files_by_extension(source_path, target_path, output_callback=None):
    """
    根据文件扩展名分类文件
    
    Args:
        source_path (str): 源文件目录路径
        target_path (str): 目标目录路径
        output_callback (function): 输出回调函数，用于将日志信息传递给GUI
    """
    # 检查源目录是否存在
    if not os.path.exists(source_path):
        error_msg = f"源目录 {source_path} 不存在"
        if output_callback:
            output_callback(error_msg)
        else:
            print(error_msg)
        raise FileNotFoundError(error_msg)
    
    # 预先收集所有文件
    source = Path(source_path)
    all_files = [f for f in source.iterdir() if f.is_file()]
    
    # 按扩展名组织文件
    extension_files = defaultdict(list)
    for file_path in all_files:
        extension = file_path.suffix.lower()
        # 如果没有扩展名，归类到"无扩展名"文件夹
        if not extension:
            extension = "无扩展名"
        extension_files[extension].append(file_path)
    
    # 处理按扩展名分类的文件
    destination_folders = set()
    moved_files_count = 0
    for extension, files in extension_files.items():
        # 创建以扩展名命名的文件夹（包括点号）
        destination_folder = os.path.join(target_path, extension)
        destination_folders.add(destination_folder)
        for file_path in files:
            success, dest_path = move_file(file_path, destination_folder, output_callback)
            if success:
                moved_files_count += 1
    
    # 输出操作总结
    summary_msg = f"按扩展名分类操作成功完成！总共处理了 {len(all_files)} 个文件，实际移动了 {moved_files_count} 个文件"
    folder_msg = "文件已移动到以下文件夹:"
    if output_callback:
        output_callback(summary_msg)
        output_callback(folder_msg)
        for folder in sorted(destination_folders):
            output_callback(f"  {folder}")
    else:
        print(summary_msg)
        print(folder_msg)
        for folder in sorted(destination_folders):
            print(f"  {folder}")
    
    return True

def get_supported_types():
    """获取所有支持的文件类型"""
    return list(FILE_TYPES.keys())

def file_classify():
    parser = argparse.ArgumentParser(description='文件自动分类工具')
    parser.add_argument('source_directory', help='源目录路径')
    parser.add_argument('-t', '--target', help='目标目录路径（默认为源目录）')
    parser.add_argument('-f', '--formats', nargs='+', help='要分类的文件格式，如：image.png_image image.jpg_image（默认为全部）')
    parser.add_argument('-n', '--name_patterns', nargs='+', help='按文件名模式分类，如：date number')
    parser.add_argument('-k', '--keywords', nargs='+', help='按关键词分类，如：project report')
    parser.add_argument('-l', '--list', action='store_true', help='列出所有支持的文件格式')
    
    args = parser.parse_args()
    
    # 如果使用了-l参数，列出所有支持的格式并退出
    if args.list:
        print("支持的文件格式:")
        for file_type, extensions in FILE_TYPES.items():
            print(f"  {file_type}: {', '.join(extensions)}")
        print("\n支持的子文件格式:")
        for main_type, sub_types in SUB_TYPES.items():
            for sub_type_name, extensions in sub_types.items():
                print(f"  {main_type}.{sub_type_name}: {', '.join(extensions)}")
        print("\n支持的文件名模式:")
        for pattern_name, pattern in NAME_PATTERNS.items():
            print(f"  {pattern_name}: {pattern}")
        return
    
    # 检查源目录是否存在
    if not os.path.exists(args.source_directory):
        print(f"错误: 源目录 '{args.source_directory}' 不存在")
        return
    
    # 处理用户输入的模式名称，转换为实际的正则表达式
    name_patterns = None
    if args.name_patterns:
        name_patterns = {}
        for pattern_name in args.name_patterns:
            if pattern_name in NAME_PATTERNS:
                name_patterns[pattern_name] = NAME_PATTERNS[pattern_name]
            else:
                print(f"警告: 未知的文件名模式 '{pattern_name}'，将忽略")
    
    # 使用参数调用分类函数
    target_directory = args.target if args.target else args.source_directory
    classify_files(
        args.source_directory, 
        target_directory, 
        args.formats if args.formats else None,
        name_patterns if name_patterns else None,
        args.keywords if args.keywords else None
    )
    print("文件分类完成!")