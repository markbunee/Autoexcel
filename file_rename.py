import os
import shutil
from pathlib import Path

def rename_files_sequentially(source_path, prefix="", start_number=1, digits=4, keyword="", extension_filter=None, output_callback=None):
    """
    对文件进行批量重命名，按顺序编号
    
    Args:
        source_path (str): 源文件目录路径
        prefix (str): 文件名前缀，默认为空
        start_number (int): 起始编号，默认为1
        digits (int): 编号位数，默认为4位（如0001）
        keyword (str): 关键词，添加在序号后，默认为空
        extension_filter (list): 文件扩展名过滤器，只重命名指定类型的文件，如['.txt', '.pdf']，默认为None表示所有文件
        output_callback (function): 输出回调函数，用于将日志信息传递给GUI
    """
    # 检查源目录是否存在
    if not os.path.exists(source_path):
        raise FileNotFoundError(f"源目录 {source_path} 不存在")
    
    # 构造关键词部分
    keyword_part = f"_{keyword}" if keyword else ""
    
    # 内部打印函数，支持GUI输出
    def _print(msg):
        """内部打印函数，支持GUI输出"""
        if output_callback:
            output_callback(msg)
        else:
            print(msg)
    
    # 获取目录中的所有文件
    source = Path(source_path)
    files = []
    
    for file_path in source.iterdir():
        if file_path.is_file():
            # 如果指定了扩展名过滤器，则只添加匹配的文件
            if extension_filter is None or file_path.suffix.lower() in [ext.lower() for ext in extension_filter]:
                files.append(file_path)
    
    # 按文件名排序，确保重命名的一致性
    files.sort()
    
    # 预先检查是否有重命名冲突
    new_names = set()
    conflicts = []
    for i, file_path in enumerate(files):
        number = start_number + i
        formatted_number = str(number).zfill(digits)
        new_filename = f"{prefix}{formatted_number}{keyword_part}{file_path.suffix}"
        if new_filename in new_names:
            conflicts.append((file_path.name, new_filename))
        new_names.add(new_filename)
    
    if conflicts:
        _print("警告：以下文件重命名会导致冲突:")
        for old_name, new_name in conflicts:
            _print(f"  {old_name} -> {new_name}")
        response = input("是否继续重命名? (y/N): ")
        if response.lower() != 'y':
            _print("重命名操作已取消")
            return False
    
    # 重命名文件
    renamed_count = 0
    for i, file_path in enumerate(files):
        # 计算新文件名
        number = start_number + i
        formatted_number = str(number).zfill(digits)
        new_filename = f"{prefix}{formatted_number}{keyword_part}{file_path.suffix}"
        new_file_path = file_path.parent / new_filename
        
        # 重命名文件
        try:
            # 检查目标文件是否已存在
            if new_file_path.exists() and new_file_path != file_path:
                _print(f"警告：文件 {new_filename} 已存在，跳过 {file_path.name}")
                continue
                
            file_path.rename(new_file_path)
            _print(f"已重命名: {file_path.name} -> {new_filename}")
            renamed_count += 1
        except Exception as e:
            _print(f"重命名文件失败 {file_path.name}: {str(e)}")
    
    _print(f"成功重命名 {renamed_count} 个文件")
    return True

def rename_files_with_keyword_pattern(source_path, keywords, prefix="", start_number=1, digits=4, extension_filter=None, output_callback=None):
    """
    根据关键词列表对文件进行重命名，关键词来自预定义列表
    
    Args:
        source_path (str): 源文件目录路径
        keywords (list): 关键词列表，如['财经报告', '政治报告', '学习报告']
        prefix (str): 文件名前缀，默认为空
        start_number (int): 起始编号，默认为1
        digits (int): 编号位数，默认为4位（如0001）
        extension_filter (list): 文件扩展名过滤器，只重命名指定类型的文件，如['.txt', '.pdf']，默认为None表示所有文件
        output_callback (function): 输出回调函数，用于将日志信息传递给GUI
    """
    def _print(msg):
        """内部打印函数，支持GUI输出"""
        if output_callback:
            output_callback(msg)
        else:
            print(msg)
    
    # 检查源目录是否存在
    if not os.path.exists(source_path):
        raise FileNotFoundError(f"源目录 {source_path} 不存在")
    
    # 获取目录中的所有文件
    source = Path(source_path)
    files = []
    
    for file_path in source.iterdir():
        if file_path.is_file():
            # 如果指定了扩展名过滤器，则只添加匹配的文件
            if extension_filter is None or file_path.suffix.lower() in [ext.lower() for ext in extension_filter]:
                files.append(file_path)
    
    # 按文件名排序，确保重命名的一致性
    files.sort()
    
    # 预先检查是否有重命名冲突
    new_names = set()
    conflicts = []
    for i, file_path in enumerate(files):
        # 确定使用的关键词
        keyword = ""
        if i < len(keywords):
            keyword = keywords[i]
        else:
            # 如果文件数超过关键词数，使用默认编号
            keyword = f"file_{i+1}"
        
        # 计算新文件名
        number = start_number + i
        formatted_number = str(number).zfill(digits)
        keyword_part = f"_{keyword}" if keyword else ""
        new_filename = f"{prefix}{formatted_number}{keyword_part}{file_path.suffix}"
        if new_filename in new_names:
            conflicts.append((file_path.name, new_filename))
        new_names.add(new_filename)
    
    if conflicts:
        _print("警告：以下文件重命名会导致冲突:")
        for old_name, new_name in conflicts:
            _print(f"  {old_name} -> {new_name}")
        response = input("是否继续重命名? (y/N): ")
        if response.lower() != 'y':
            _print("重命名操作已取消")
            return False
    
    # 重命名文件
    renamed_count = 0
    for i, file_path in enumerate(files):
        # 确定使用的关键词
        keyword = ""
        if i < len(keywords):
            keyword = keywords[i]
        else:
            # 如果文件数超过关键词数，使用默认编号
            keyword = f"file_{i+1}"
        
        # 计算新文件名
        number = start_number + i
        formatted_number = str(number).zfill(digits)
        keyword_part = f"_{keyword}" if keyword else ""
        new_filename = f"{prefix}{formatted_number}{keyword_part}{file_path.suffix}"
        new_file_path = file_path.parent / new_filename
        
        # 重命名文件
        try:
            # 检查目标文件是否已存在
            if new_file_path.exists() and new_file_path != file_path:
                _print(f"警告：文件 {new_filename} 已存在，跳过 {file_path.name}")
                continue
                
            file_path.rename(new_file_path)
            _print(f"已重命名: {file_path.name} -> {new_filename}")
            renamed_count += 1
        except Exception as e:
            _print(f"重命名文件失败 {file_path.name}: {str(e)}")
    
    _print(f"成功重命名 {renamed_count} 个文件")
    return True

def rename_files_extract_keyword(source_path, prefix="", start_number=1, digits=4, keyword_patterns=None, extension_filter=None, output_callback=None):
    """
    从文件名中提取关键词并用于重命名
    
    Args:
        source_path (str): 源文件目录路径
        prefix (str): 文件名前缀，默认为空
        start_number (int): 起始编号，默认为1
        digits (int): 编号位数，默认为4位（如0001）
        keyword_patterns (list): 关键词模式列表，用于从文件名中提取关键词
        extension_filter (list): 文件扩展名过滤器，只重命名指定类型的文件，如['.txt', '.pdf']，默认为None表示所有文件
        output_callback (function): 输出回调函数，用于将日志信息传递给GUI
    """
    def _print(msg):
        """内部打印函数，支持GUI输出"""
        if output_callback:
            output_callback(msg)
        else:
            print(msg)
    
    # 检查源目录是否存在
    if not os.path.exists(source_path):
        raise FileNotFoundError(f"源目录 {source_path} 不存在")
    
    # 获取目录中的所有文件
    source = Path(source_path)
    files = []
    
    for file_path in source.iterdir():
        if file_path.is_file():
            # 如果指定了扩展名过滤器，则只添加匹配的文件
            if extension_filter is None or file_path.suffix.lower() in [ext.lower() for ext in extension_filter]:
                files.append(file_path)
    
    # 按文件名排序，确保重命名的一致性
    files.sort()
    
    # 预先检查是否有重命名冲突
    new_names = set()
    conflicts = []
    for i, file_path in enumerate(files):
        # 从文件名中提取关键词
        keyword = ""
        filename_without_ext = file_path.stem
        
        if keyword_patterns:
            for pattern in keyword_patterns:
                if pattern in filename_without_ext:
                    keyword = pattern
                    break
        
        # 如果没有匹配的关键词，则使用文件名的一部分
        if not keyword:
            keyword = filename_without_ext
        
        # 计算新文件名
        number = start_number + i
        formatted_number = str(number).zfill(digits)
        keyword_part = f"_{keyword}" if keyword else ""
        new_filename = f"{prefix}{formatted_number}{keyword_part}{file_path.suffix}"
        if new_filename in new_names:
            conflicts.append((file_path.name, new_filename))
        new_names.add(new_filename)
    
    if conflicts:
        _print("警告：以下文件重命名会导致冲突:")
        for old_name, new_name in conflicts:
            _print(f"  {old_name} -> {new_name}")
        response = input("是否继续重命名? (y/N): ")
        if response.lower() != 'y':
            _print("重命名操作已取消")
            return False
    
    # 重命名文件
    renamed_count = 0
    for i, file_path in enumerate(files):
        # 从文件名中提取关键词
        keyword = ""
        filename_without_ext = file_path.stem
        
        if keyword_patterns:
            for pattern in keyword_patterns:
                if pattern in filename_without_ext:
                    keyword = pattern
                    break
        
        # 如果没有匹配的关键词，则使用文件名的一部分
        if not keyword:
            keyword = filename_without_ext
        
        # 计算新文件名
        number = start_number + i
        formatted_number = str(number).zfill(digits)
        keyword_part = f"_{keyword}" if keyword else ""
        new_filename = f"{prefix}{formatted_number}{keyword_part}{file_path.suffix}"
        new_file_path = file_path.parent / new_filename
        
        # 重命名文件
        try:
            # 检查目标文件是否已存在
            if new_file_path.exists() and new_file_path != file_path:
                _print(f"警告：文件 {new_filename} 已存在，跳过 {file_path.name}")
                continue
                
            file_path.rename(new_file_path)
            _print(f"已重命名: {file_path.name} -> {new_filename}")
            renamed_count += 1
        except Exception as e:
            _print(f"重命名文件失败 {file_path.name}: {str(e)}")
    
    _print(f"成功重命名 {renamed_count} 个文件")
    return True

# 示例用法
if __name__ == "__main__":
    # 示例1: 基础顺序重命名，带关键词
    # rename_files_sequentially("./test_files", prefix="", digits=4, keyword="政治报告")
    
    # 示例2: 根据关键词列表重命名
    # keywords = ['财经报告', '政治报告', '学习报告', '美国政治报告', '虚拟经济政治报告', '外交政治与协作']
    # rename_files_with_keyword_pattern("./test_files", keywords, digits=4)
    
    # 示例3: 从文件名中提取关键词
    # keyword_patterns = ['财经', '政治', '学习', '美国', '虚拟经济', '外交']
    # rename_files_extract_keyword("./test_files", digits=4, keyword_patterns=keyword_patterns)
    pass