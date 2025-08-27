import pandas as pd
import os
import difflib
#一致性评价进度excel专用代码批量0.4文本相似度批量本地python库筛选

def calculate_text_similarity(text1, text2):
    """计算两个文本的相似度分数（0-1）"""
    return difflib.SequenceMatcher(None, text1, text2).ratio()

def remove_last_two_chars(text):
    """去掉文本的最后两个字（如果长度允许）"""
    if len(text) >= 2:
        return text[:-2]
    return text

def process_excel(input_path, output_folder, similarity_threshold=0.4, anchor_column=5, 
                  compare_column=3, group_column=0, trim_chars=2, output_callback=None):
    """
    处理Excel文件，根据相似度匹配行
    
    参数:
    input_path: 输入文件路径
    output_folder: 输出文件夹路径
    similarity_threshold: 相似度阈值，默认0.4
    anchor_column: 锚点列索引，默认5（第6列）
    compare_column: 比较列索引，默认3（第4列）
    group_column: 分组列索引，默认0（第1列）
    trim_chars: 去掉末尾字符数，默认2
    output_callback: 输出回调函数，用于GUI界面显示日志
    """
    
    def print_log(message):
        """日志输出函数"""
        if output_callback:
            output_callback(message)
        else:
            print(message)
    
    # 获取原始文件名（不含路径）
    original_filename = os.path.basename(input_path)
    
    # 读取Excel文件（不自动识别表头）
    df = pd.read_excel(input_path, header=None)
    
    # 存储匹配行的索引和结果数据
    matched_indices = set()
    matched_rows = []
    
    # 创建比较列文本列表用于快速比较
    col_texts = [str(text) if not pd.isna(text) else "" for text in df[compare_column]]
    
    # 用于记录哪些锚点行找到了匹配项
    anchor_with_matches = set()
    
    # 遍历每一行（锚点行）
    for idx, row in df.iterrows():
        # 检查锚点列是否非空
        if not pd.isna(row[anchor_column]):
            group_value = row[group_column]  # 分组列的值
            compare_text = str(row[compare_column]) if not pd.isna(row[compare_column]) else ""  # 比较列的值
            
            # 去掉锚点文本的最后指定字符数
            if len(compare_text) >= trim_chars:
                processed_compare_text = compare_text[:-trim_chars]
            else:
                processed_compare_text = compare_text
            
            # 标记是否找到匹配项
            found_match = False
            
            # 全局检索匹配条件的行
            for sub_idx, sub_row in df.iterrows():
                # 跳过自身
                if idx == sub_idx:
                    continue
                    
                sub_group_value = sub_row[group_column]  # 目标行分组列的值
                sub_text = col_texts[sub_idx]  # 目标行比较列的文本
                
                # 去掉目标文本的最后指定字符数
                if len(sub_text) >= trim_chars:
                    processed_sub_text = sub_text[:-trim_chars]
                else:
                    processed_sub_text = sub_text
                
                # 双重条件匹配：
                # 1. 分组列值必须相等
                # 2. 比较列文本相似度达到阈值
                if (sub_group_value == group_value) and (calculate_text_similarity(processed_compare_text, processed_sub_text) > similarity_threshold):
                    matched_indices.add(sub_idx)
                    matched_rows.append(sub_row.tolist())
                    found_match = True
            
            # 如果找到匹配项，添加锚点行本身
            if found_match:
                matched_indices.add(idx)
                matched_rows.append(row.tolist())
                anchor_with_matches.add(idx)
    
    # 统计未找到匹配的锚点行
    all_anchors = [idx for idx, row in df.iterrows() if not pd.isna(row[anchor_column])]
    anchors_without_matches = set(all_anchors) - anchor_with_matches
    
    # 创建结果DataFrame
    result_df = pd.DataFrame(matched_rows)
    
    # 确保输出目录存在
    os.makedirs(output_folder, exist_ok=True)
    
    # 使用原始文件名命名结果文件
    output_path = os.path.join(output_folder, original_filename)
    
    # 保存结果到Excel
    result_df.to_excel(output_path, index=False, header=False)
    
    # 输出处理信息
    print_log(f"\n处理文件: {original_filename}")
    print_log(f"总行数: {len(df)}")
    print_log(f"匹配行数: {len(matched_rows)}")
    
    # 输出匹配行的原始位置（1-based）
    sorted_indices = sorted(matched_indices)
    if sorted_indices:
        positions = ", ".join([str(idx + 1) for idx in sorted_indices])
        print_log(f"匹配行在原始文件中的位置: {positions}")
    else:
        print_log("未找到匹配的行")
    
    # 输出未找到匹配的锚点行位置
    if anchors_without_matches:
        anchor_positions = ", ".join([str(idx + 1) for idx in sorted(anchors_without_matches)])
        print_log(f"未找到匹配项的锚点行位置: {anchor_positions}")
    else:
        print_log("所有锚点行都找到了匹配项")
    
    print_log(f"结果已保存到: {output_path}")
    
    return output_path

# 使用示例
if __name__ == "__main__":
    # 定义输入文件列表
    input_files = [
        # r"D:\ASUS\working\一致性评价进度\一致性评价进度\一致性评价进度-感觉系统药物.xlsx",
        # r"D:\ASUS\working\一致性评价进度\一致性评价进度\一致性评价进度-呼吸系统疾病用药.xlsx",
        # r"D:\ASUS\working\一致性评价进度\一致性评价进度\一致性评价进度-呼吸系统用药.xlsx",
        
        r"D:\ASUS\working\一致性评价进度\一致性评价进度\一致性评价进度-肌肉-骨骼系统.xlsx",
        # r"D:\ASUS\working\一致性评价进度\一致性评价进度\一致性评价进度-抗寄生虫药、杀虫剂和驱虫剂.xlsx",
        # r"D:\ASUS\working\一致性评价进度\一致性评价进度\一致性评价进度-抗肿瘤和免疫调节剂.xlsx",
        # r"D:\ASUS\working\一致性评价进度\一致性评价进度\一致性评价进度-皮肤病用药.xlsx",
        # r"D:\ASUS\working\一致性评价进度\一致性评价进度\一致性评价进度-全身用激素类制剂(不含性激素和胰岛素).xlsx",
        # r"D:\ASUS\working\一致性评价进度\一致性评价进度\一致性评价进度-全身用抗感染药物.xlsx",
        # r"D:\ASUS\working\一致性评价进度\一致性评价进度\一致性评价进度-神经系统药物.xlsx",
        # r"D:\ASUS\working\一致性评价进度\一致性评价进度\一致性评价进度-生殖泌尿系统和性激素类药物.xlsx",
        # r"D:\ASUS\working\一致性评价进度\一致性评价进度\一致性评价进度-未分类.xlsx",
        # r"D:\ASUS\working\一致性评价进度\一致性评价进度\一致性评价进度-消化系统及代谢药.xlsx",
        # r"D:\ASUS\working\一致性评价进度\一致性评价进度\一致性评价进度-心脑血管系统药物.xlsx",
        # r"D:\ASUS\working\一致性评价进度\一致性评价进度\一致性评价进度-血液和造血系统药物.xlsx",
        # r"D:\ASUS\working\一致性评价进度\一致性评价进度\一致性评价进度-原料药及非直接作用于人体药物.xlsx",
        # r"D:\ASUS\working\一致性评价进度\一致性评价进度\一致性评价进度-杂类.xlsx",

    ]
    
    # 输出目录
    output_dir = r"D:\ASUS\working\output"
    
    # 批量处理所有文件
    for input_file in input_files:
        # 确保文件存在
        if os.path.exists(input_file):
            # 处理当前文件
            result_path = process_excel(input_file, output_dir)
        else:
            print(f"\n文件不存在: {input_file}")