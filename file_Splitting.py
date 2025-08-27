import pandas as pd
import os
from pathlib import Path

def split_excel_by_column(source_file, split_column, output_columns, output_dir="split_results", output_callback=None):
    """
    根据指定列分割Excel文件为多个文件
    
    Args:
        source_file (str): 源Excel文件路径
        split_column (int): 用于分割的列序号（从1开始数起）
        output_columns (list): 要输出的列组合列表，每个组合是一个列序号列表（从1开始数起）
        output_dir (str): 输出目录路径，默认为"split_results"
        output_callback (function): 输出回调函数，用于将日志信息传递给GUI
        
    Returns:
        bool: 分割是否成功
    """
    def _print(msg):
        """内部打印函数，支持GUI输出"""
        if output_callback:
            output_callback(msg)
        else:
            print(msg)
    
    try:
        # 检查源文件是否存在
        if not os.path.exists(source_file):
            raise FileNotFoundError(f"源文件不存在: {source_file}")
        
        # 创建输出目录
        os.makedirs(output_dir, exist_ok=True)
        
        # 读取Excel文件
        df = pd.read_excel(source_file, dtype=str)  # 使用字符串类型避免类型转换问题
        
        # 检查列索引是否有效
        if split_column > len(df.columns) or split_column < 1:
            raise ValueError(f"分割列索引 {split_column} 超出范围，文件共有 {len(df.columns)} 列")
        
        # 转换为从0开始的索引
        split_col_index = split_column - 1
        
        # 获取分割列的唯一值
        unique_values = df.iloc[:, split_col_index].dropna().unique()
        
        _print(f"根据列 '{df.columns[split_col_index]}' 分割，共有 {len(unique_values)} 个唯一值")
        
        # 为每个唯一值创建一个文件
        for i, value in enumerate(unique_values):
            # 筛选数据
            filtered_df = df[df.iloc[:, split_col_index] == value]
            
            if len(filtered_df) == 0:
                continue
            
            # 如果未指定输出列，则输出所有列
            if output_columns is None:
                # 生成文件名
                safe_value = str(value).replace('/', '_').replace('\\', '_').replace(':', '_').replace('*', '_').replace('?', '_').replace('"', '_').replace('<', '_').replace('>', '_').replace('|', '_')
                filename = f"{safe_value}.xlsx"
                filepath = os.path.join(output_dir, filename)
                
                # 保存文件
                filtered_df.to_excel(filepath, index=False)
                _print(f"已保存: {filepath}")
            else:
                # 为每个输出列组合创建文件
                for j, columns in enumerate(output_columns):
                    # 转换列索引为从0开始
                    zero_based_columns = [col - 1 for col in columns]
                    
                    # 检查列索引是否有效
                    invalid_cols = [col for col in zero_based_columns if col >= len(df.columns) or col < 0]
                    if invalid_cols:
                        _print(f"警告: 列索引 {invalid_cols} 超出范围，跳过该列组合")
                        continue
                    
                    # 选择指定列
                    selected_df = filtered_df.iloc[:, zero_based_columns]
                    
                    # 生成文件名
                    safe_value = str(value).replace('/', '_').replace('\\', '_').replace(':', '_').replace('*', '_').replace('?', '_').replace('"', '_').replace('<', '_').replace('>', '_').replace('|', '_')
                    filename = f"{safe_value}_{j+1}.xlsx"
                    filepath = os.path.join(output_dir, filename)
                    
                    # 保存文件
                    selected_df.to_excel(filepath, index=False)
                    _print(f"已保存: {filepath}")
        
        _print(f"分割完成，结果保存在目录: {output_dir}")
        return True
        
    except Exception as e:
        _print(f"分割文件时出错: {str(e)}")
        return False

def split_excel_by_column_advanced(source_file, split_configs, output_dir="split_results", output_callback=None):
    """
    高级分割功能，支持多种分割配置
    
    Args:
        source_file (str): 源Excel文件路径
        split_configs (list): 分割配置列表，每个配置是一个字典，包含:
            - 'split_column': 用于分割的列序号（从1开始数起）
            - 'output_columns': 要输出的列序号列表（从1开始数起）
            - 'output_name': 输出文件名前缀（可选）
        output_dir (str): 输出目录路径，默认为"split_results"
        output_callback (function): 输出回调函数，用于将日志信息传递给GUI
        
    Returns:
        bool: 分割是否成功
    """
    def _print(msg):
        """内部打印函数，支持GUI输出"""
        if output_callback:
            output_callback(msg)
        else:
            print(msg)
    
    try:
        # 检查源文件是否存在
        if not os.path.exists(source_file):
            raise FileNotFoundError(f"源文件不存在: {source_file}")
        
        # 创建输出目录
        os.makedirs(output_dir, exist_ok=True)
        
        # 读取Excel文件
        df = pd.read_excel(source_file, dtype=str)  # 使用字符串类型避免类型转换问题
        
        # 处理每个分割配置
        for config_idx, config in enumerate(split_configs):
            split_column = config['split_column']
            output_columns = config['output_columns']
            output_name = config.get('output_name', f'split_{config_idx+1}')
            
            # 检查列索引是否有效
            if split_column > len(df.columns) or split_column < 1:
                _print(f"警告: 分割列索引 {split_column} 超出范围，跳过该配置")
                continue
                
            # 转换为从0开始的索引
            split_col_index = split_column - 1
            
            # 获取分割列的唯一值
            unique_values = df.iloc[:, split_col_index].dropna().unique()
            
            _print(f"配置 {config_idx+1}: 根据列 '{df.columns[split_col_index]}' 分割，共有 {len(unique_values)} 个唯一值")
            
            # 为每个唯一值创建一个文件
            for value in unique_values:
                # 筛选数据
                filtered_df = df[df.iloc[:, split_col_index] == value]
                
                if len(filtered_df) == 0:
                    continue
                
                # 转换列索引为从0开始
                zero_based_columns = [col - 1 for col in output_columns]
                
                # 检查列索引是否有效
                invalid_cols = [col for col in zero_based_columns if col >= len(df.columns) or col < 0]
                if invalid_cols:
                    _print(f"警告: 列索引 {invalid_cols} 超出范围，跳过该列组合")
                    continue
                
                # 选择指定列
                selected_df = filtered_df.iloc[:, zero_based_columns]
                
                # 生成文件名
                safe_value = str(value).replace('/', '_').replace('\\', '_').replace(':', '_').replace('*', '_').replace('?', '_').replace('"', '_').replace('<', '_').replace('>', '_').replace('|', '_')
                filename = f"{output_name}_{safe_value}.xlsx"
                filepath = os.path.join(output_dir, filename)
                
                # 保存文件
                selected_df.to_excel(filepath, index=False)
                _print(f"已保存: {filepath}")
        
        _print(f"高级分割完成，结果保存在目录: {output_dir}")
        return True
        
    except Exception as e:
        _print(f"分割文件时出错: {str(e)}")
        return False

# 示例用法
if __name__ == "__main__":
    # 示例1: 基本分割
    # split_excel_by_column(
    #     source_file="data.xlsx",
    #     split_column=1,  # 按第1列（企业名）分割
    #     output_columns=[
    #         [1, 2],  # 输出第1列和第2列（企业名和销售额）
    #         [1, 3],  # 输出第1列和第3列（企业名和销售量）
    #         [1, 4, 5]  # 输出第1列、第4列和第5列（企业名、一致性评价进度和国家）
    #     ],
    #     output_dir="split_results"
    # )
    
    # 示例2: 高级分割
    # split_configs = [
    #     {
    #         'split_column': 1,  # 按第1列（企业名）分割
    #         'output_columns': [1, 2],  # 输出第1列和第2列
    #         'output_name': '企业销售额'
    #     },
    #     {
    #         'split_column': 1,  # 按第1列（企业名）分割
    #         'output_columns': [1, 3],  # 输出第1列和第3列
    #         'output_name': '企业销售量'
    #     }
    # ]
    # split_excel_by_column_advanced("data.xlsx", split_configs, "advanced_split_results")
    pass