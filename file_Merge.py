import pandas as pd
import os
from pathlib import Path

def merge_excel_files_by_column(file_paths, match_columns, target_columns=None, output_path="merged_result.xlsx", chunk_size=10000, output_callback=None):
    """
    根据指定列匹配多份Excel文件并合并到一个Excel文件中
    
    Args:
        file_paths (list): Excel文件路径列表
        match_columns (list): 用于匹配的列序号列表（从1开始数起）
        target_columns (list): 每个文件要合并的列序号列表（从1开始数起），None表示合并所有列
        output_path (str): 输出文件路径，默认为"merged_result.xlsx"
        chunk_size (int): 分块处理大小，默认为10000行
        output_callback (function): 输出回调函数，用于将日志信息传递给GUI
        
    Returns:
        bool: 合并是否成功
    """
    def _print(msg):
        """内部打印函数，支持GUI输出"""
        if output_callback:
            output_callback(msg)
        else:
            print(msg)
    
    if len(file_paths) != len(match_columns):
        error_msg = "文件路径数量与匹配列数量不一致"
        _print(error_msg)
        raise ValueError(error_msg)
    
    if target_columns is not None and len(file_paths) != len(target_columns):
        error_msg = "文件路径数量与目标列数量不一致"
        _print(error_msg)
        raise ValueError(error_msg)
    
    try:
        # 检查文件是否存在
        for file_path in file_paths:
            if not os.path.exists(file_path):
                error_msg = f"文件不存在: {file_path}"
                _print(error_msg)
                raise FileNotFoundError(error_msg)
        
        # 读取第一个文件作为基础
        # 转换为从0开始的索引
        match_col_index = match_columns[0] - 1
        target_cols = None if target_columns is None else [col - 1 for col in target_columns[0]]
        
        # 读取第一个Excel文件
        df_main = pd.read_excel(file_paths[0], dtype=str)  # 使用字符串类型避免类型转换问题
        
        # 获取匹配列的名称
        match_column_name = df_main.columns[match_col_index]
        
        # 如果指定了目标列，则只选择这些列
        if target_cols is not None:
            # 确保匹配列在选择的列中
            if match_col_index not in target_cols:
                target_cols.append(match_col_index)
            df_main = df_main.iloc[:, sorted(target_cols)]
        else:
            # 如果未指定目标列，则选择所有列
            target_cols = list(range(len(df_main.columns)))
        
        # 重命名列，添加文件标识
        new_columns = []
        for i, col in enumerate(df_main.columns):
            if i == match_col_index:
                new_columns.append(col)  # 匹配列保持原名
            else:
                new_columns.append(f"{col}_file1")
        df_main.columns = new_columns
        
        # 依次处理其他文件
        for file_idx, (file_path, match_col) in enumerate(zip(file_paths[1:], match_columns[1:]), start=2):
            _print(f"正在处理文件: {file_path}")
            
            # 转换为从0开始的索引
            match_col_index_other = match_col - 1
            target_cols_other = None if target_columns is None else [col - 1 for col in target_columns[file_idx-1]]
            
            # 读取Excel文件
            df_other = pd.read_excel(file_path, dtype=str)
            
            # 获取匹配列的名称
            match_column_name_other = df_other.columns[match_col_index_other]
            
            # 如果指定了目标列，则只选择这些列
            if target_cols_other is not None:
                # 确保匹配列在选择的列中
                if match_col_index_other not in target_cols_other:
                    target_cols_other.append(match_col_index_other)
                df_other = df_other.iloc[:, sorted(target_cols_other)]
            else:
                # 如果未指定目标列，则选择所有列
                target_cols_other = list(range(len(df_other.columns)))
            
            # 重命名列，添加文件标识
            new_columns_other = []
            for i, col in enumerate(df_other.columns):
                if i == match_col_index_other:
                    new_columns_other.append(col)  # 匹配列保持原名
                else:
                    new_columns_other.append(f"{col}_file{file_idx}")
            df_other.columns = new_columns_other
            
            # 根据匹配列合并数据
            # 使用第一个文件的匹配列名作为连接键
            df_main = pd.merge(df_main, df_other, left_on=match_column_name, right_on=match_column_name_other, how='outer')
            
            # 删除重复的匹配列
            if match_column_name != match_column_name_other:
                df_main.drop(columns=[match_column_name_other], inplace=True)
            
            # 释放内存
            del df_other
        
        # 保存合并后的数据
        df_main.to_excel(output_path, index=False)
        _print(f"文件已成功合并并保存到: {output_path}")
        
        return True
        
    except Exception as e:
        error_msg = f"合并文件时出错: {str(e)}"
        _print(error_msg)
        return False

def merge_excel_files_simple(file_paths, output_path="merged_result.xlsx", chunk_size=10000):
    """
    简单合并多份Excel文件（按行合并，不基于列匹配）
    
    Args:
        file_paths (list): Excel文件路径列表
        output_path (str): 输出文件路径，默认为"merged_result.xlsx"
        chunk_size (int): 分块处理大小，默认为10000行（此参数在当前实现中未使用）
        
    Returns:
        bool: 合并是否成功
    """
    try:
        # 检查文件是否存在
        for file_path in file_paths:
            if not os.path.exists(file_path):
                raise FileNotFoundError(f"文件不存在: {file_path}")
        
        # 收集所有数据
        all_data = []
        for file_idx, file_path in enumerate(file_paths):
            print(f"正在处理文件: {file_path}")
            
            # 读取文件
            df = pd.read_excel(file_path, dtype=str)
            df['source_file'] = f"file_{file_idx+1}"  # 添加来源文件标识
            all_data.append(df)
        
        # 合并所有数据
        if all_data:
            merged_df = pd.concat(all_data, ignore_index=True)
            # 保存合并后的数据
            merged_df.to_excel(output_path, index=False)
            print(f"文件已成功合并并保存到: {output_path}")
            print(f"合并后数据总行数: {len(merged_df)}")
            return True
        else:
            print("没有数据可以合并")
            return False
        
    except Exception as e:
        print(f"合并文件时出错: {str(e)}")
        return False

# 示例用法
if __name__ == "__main__":
    # 示例1: 根据指定列匹配合并
    # file_paths = ["file1.xlsx", "file2.xlsx", "file3.xlsx"]
    # match_columns = [1, 2, 1]  # 第1个文件的第1列，第2个文件的第2列，第3个文件的第1列作为匹配列
    # target_columns = [[1, 2, 3], [1, 2, 4], [1, 3]]  # 每个文件要合并的列
    # merge_excel_files_by_column(file_paths, match_columns, target_columns, "output.xlsx")
    
    # 示例2: 合并所有列，基于指定匹配列
    # file_paths = ["file1.xlsx", "file2.xlsx"]
    # match_columns = [1, 1]  # 两个文件都使用第1列作为匹配列
    # merge_excel_files_by_column(file_paths, match_columns, output_path="output.xlsx")
    
    # 示例3: 简单合并（按行）
    # file_paths = ["file1.xlsx", "file2.xlsx", "file3.xlsx"]
    # merge_excel_files_simple(file_paths, "simple_merged.xlsx")
    pass