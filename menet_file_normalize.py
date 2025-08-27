import pandas as pd
import os
import numpy as np
from collections import OrderedDict
import rapidfuzz
from rapidfuzz import fuzz, distance
import re
import time

def process_indication_standardization(input_file, output_folder, column_index=3, group_column_index=1, 
                                      similarity_threshold=85, edit_distance_threshold=3, min_text_length=4,
                                      output_callback=None):
    """
    适应症写法规范化处理函数
    
    参数:
    input_file: 输入Excel文件路径
    output_folder: 输出文件夹路径
    column_index: 需要处理的列索引（默认为3，即第4列）
    group_column_index: 分组列索引（默认为1，即第2列）
    similarity_threshold: 相似度阈值（默认85）
    edit_distance_threshold: 编辑距离阈值（默认3）
    min_text_length: 最小文本长度（默认4）
    output_callback: 输出回调函数，用于GUI界面显示日志
    """
    
    # 设置参数
    SIMILARITY_THRESHOLD = similarity_threshold
    EDIT_DISTANCE_THRESHOLD = edit_distance_threshold
    MIN_TEXT_LENGTH = min_text_length
    
    # 确保输出目录存在
    os.makedirs(output_folder, exist_ok=True)
    
    def print_log(message):
        """日志输出函数"""
        if output_callback:
            output_callback(message)
        else:
            print(message)
    
    def preprocess_text(text):
        """只保留文字内容，移除所有标点、空格和换行符"""
        if pd.isna(text):
            return ""
        
        text = str(text)
        
        # 移除特定的干扰字符，如"(1)"、"(1) 1."、"①"等
        # 移除括号内的数字，如(1)、(2)等
        text = re.sub(r'\(\d+\)', '', text)
        # 移除带点的数字，如1.、2.等
        text = re.sub(r'\d+\.', '', text)
        # 移除带圈的数字，如①、②等
        text = re.sub(r'[\u2460-\u2468]', '', text)  # ①-⑨
        text = re.sub(r'[\u3251-\u325f]', '', text)  # ㉑-㉟ (部分)
        text = re.sub(r'[\u32b1-\u32bf]', '', text)  # ㊱-㊿ (部分)
        
        # 移除其他常见的编号和列表符号
        text = re.sub(r'[\u2022\u2023\u25E6]', '', text)  # 各种项目符号：•、‣、◦
        text = re.sub(r'[\u25A0\u25A1\u25A2\u25A3\u25A4\u25A5\u25A6\u25A7\u25A8\u25A9]', '', text)  # 各种方块符号
        text = re.sub(r'[\u25B6\u25B7\u25B8\u25B9\u25BA\u25BB\u25BC\u25BD\u25BE\u25BF]', '', text)  # 各种三角符号
        text = re.sub(r'[\u25C0\u25C1\u25C2\u25C3\u25C4\u25C5\u25C6\u25C7\u25C8\u25C9]', '', text)  # 各种反向三角符号
        text = re.sub(r'[\u25CB\u25CC\u25CD\u25CE\u25CF\u25D0\u25D1\u25D2\u25D3\u25D4]', '', text)  # 各种圆形符号
        text = re.sub(r'[\u25D5\u25D6\u25D7\u25D8\u25D9\u25DA\u25DB\u25DC\u25DD\u25DE]', '', text)  # 更多圆形符号
        text = re.sub(r'[\u25DF\u25E0\u25E1\u25E2\u25E3\u25E4\u25E5]', '', text)  # 更多几何符号
        
        # 移除罗马数字（小写）
        text = re.sub(r'[\u2170-\u217f]', '', text)
        # 移除罗马数字（大写）
        text = re.sub(r'[\u2160-\u216f]', '', text)
        
        # 移除字母编号，如a)、b)、A)、B)等
        text = re.sub(r'[a-zA-Z]\)', '', text)
        
        # 移除其他常见的标点和符号
        text = re.sub(r'[\u2010-\u2015]', '', text)  # 各种连字符
        text = re.sub(r'[\u2018-\u201f]', '', text)  # 各种引号
        text = re.sub(r'[\u2026]', '', text)  # 省略号
        text = re.sub(r'[\u2032-\u2037]', '', text)  # 撇号和角分符号
        text = re.sub(r'[\u203B]', '', text)  # 参考标记符号 ※
        text = re.sub(r'[\u203C-\u203F]', '', text)  # 更多符号
        text = re.sub(r'[\u2041-\u2044]', '', text)  # 更多符号
        text = re.sub(r'[\u2047-\u2051]', '', text)  # 更多符号
        text = re.sub(r'[\u2053-\u205E]', '', text)  # 更多符号
        text = re.sub(r'[\u2060-\u2064]', '', text)  # 更多符号
        
        # 移除所有非文字字符（只保留中文、英文、数字）
        text = re.sub(r'[^\w\u4e00-\u9fff]', '', text)
        
        return text

    def group_normalize(df, col_group=1, col_target=3):
        """
        在指定分组列内容一致的组内，对目标列进行文本相似度归一化
        只匹配文字内容，忽略标点、空格和换行符的差异
        """
        start_time = time.time()
        print_log(f"开始分组文本归一化处理，分组列为: '{df.columns[col_group]}'")
        grouped_maps = {}  # 存储每个组的映射关系
        total_groups = df.iloc[:, col_group].nunique()
        processed_groups = 0
        preprocess_cache = {}  # 文本预处理结果缓存
        total_mappings = 0  # 记录总归一化条目数
        
        # 遍历每个分组
        for group_key, group_df in df.groupby(df.columns[col_group]):
            processed_groups += 1
            
            # 处理空组
            if group_key is None or pd.isna(group_key):
                continue
                
            # 进度显示
            if processed_groups % 50 == 0 or processed_groups == total_groups:
                pass
            
            # 获取组内所有行（保持原始顺序）
            group_rows = group_df.copy()
            
            # 存储每个文本的行号（用于确定标准词）
            text_to_index = {}
            for idx, row in group_rows.iterrows():
                text = row.iloc[col_target]
                if not pd.isna(text):
                    text_to_index[text] = idx
            
            # 构建组内归一化映射
            group_map = {}
            standards = []  # 存储组内的标准词列表
            
            # 遍历组内所有行
            for idx, row in group_rows.iterrows():
                text = row.iloc[col_target]
                
                # 跳过空值
                if pd.isna(text):
                    group_map[text] = text
                    continue
                
                # 短文本直接保留
                if len(str(text)) < MIN_TEXT_LENGTH:
                    group_map[text] = text
                    continue
                
                # 预处理当前文本（只保留文字内容）
                if text not in preprocess_cache:
                    preprocess_cache[text] = preprocess_text(text)
                
                # 检查是否已有匹配的标准词
                matched_standard = None
                for std in standards:
                    # 预处理标准词（只保留文字内容）
                    if std not in preprocess_cache:
                        preprocess_cache[std] = preprocess_text(std)
                    
                    # 计算相似度（只基于文字内容）
                    sim_score = fuzz.token_sort_ratio(
                        preprocess_cache[text],
                        preprocess_cache[std]
                    )
                    
                    # 计算编辑距离（只基于文字内容）
                    edit_dist = distance.Levenshtein.distance(
                        preprocess_cache[text],
                        preprocess_cache[std]
                    )
                    
                    # 检查是否满足阈值条件
                    if sim_score > SIMILARITY_THRESHOLD and edit_dist <= EDIT_DISTANCE_THRESHOLD:
                        matched_standard = std
                        break
                
                # 处理匹配结果
                if matched_standard:
                    group_map[text] = matched_standard
                    total_mappings += 1
                else:
                    # 没有匹配项，作为新标准词
                    group_map[text] = text
                    standards.append(text)
            
            grouped_maps[group_key] = group_map
        
        print_log(f"\n分组处理完成! 共处理 {len(grouped_maps)} 个分组")
        print_log(f"总共创建 {total_mappings} 条归一化映射")
        print_log(f"总耗时: {time.time()-start_time:.1f}秒")
        return grouped_maps

    # 读取Excel文件
    try:
        start_time = time.time()
        print_log(f"开始读取文件: {input_file}")
        df = pd.read_excel(input_file)
        print_log(f"成功读取文件! 共 {len(df)} 行数据 | 耗时: {time.time()-start_time:.1f}秒")
        col_name = df.columns[column_index]
        print_log(f"分析列: '{col_name}'（第{column_index+1}列）")
        
        # 创建原始列副本（用于后续归一化处理）
        original_column = df.iloc[:, column_index].copy()
    except Exception as e:
        print_log(f"读取文件失败: {str(e)}")
        return False

    # 获取唯一类别及第一条记录
    start_time = time.time()
    print_log("\n开始提取唯一类别...")
    unique_categories = OrderedDict()
    first_occurrences = []
    duplicate_count = 0

    for idx, row in df.iterrows():
        category = row.iloc[column_index]
        
        # 跳过空值
        if pd.isna(category):
            duplicate_count += 1
            continue
        
        # 如果类别未记录，保存该行数据
        if category not in unique_categories:
            unique_categories[category] = (row.copy(), idx)  # 存储行数据和行号
            first_occurrences.append(row.copy())
        else:
            duplicate_count += 1

    # 统计信息
    category_count = len(unique_categories)
    print_log(f"提取完成! 发现 {category_count} 个唯一类别 | 重复值: {duplicate_count}")
    print_log(f"处理耗时: {time.time()-start_time:.1f}秒")

    # 创建结果DataFrame
    result_df = pd.DataFrame(first_occurrences)

    # 添加说明列
    result_df.insert(0, '说明', '首次出现记录')

    # 保存结果
    output_file = os.path.join(output_folder, '类别汇总.xlsx')
    try:
        result_df.to_excel(output_file, index=False)
        print_log(f"\n结果已保存到: {output_file}")
        print_log(f"文件包含 {len(result_df)} 条记录（每个类别一条）")
        
        # 类别统计信息
        summary = f"""类别分析报告
================================
分析文件: {os.path.basename(input_file)}
分析列名: '{col_name}'（第{column_index+1}列）
分析时间: {pd.Timestamp.now()}

统计结果:
◦ 总数据行数: {len(df)}
◦ 发现唯一类别数: {category_count}
◦ 重复条目数: {duplicate_count}
◦ 输出文件路径: {output_file}

处理详情:
◦ 分组归一化列: '{df.columns[group_column_index]}'
◦ 相似度阈值: {SIMILARITY_THRESHOLD}%
◦ 编辑距离阈值: {EDIT_DISTANCE_THRESHOLD}
◦ 最小文本长度: {MIN_TEXT_LENGTH}
"""
        print_log(summary)
        
        # 保存统计报告
        report_path = os.path.join(output_folder, '类别统计报告.txt')
        with open(report_path, 'w', encoding='utf-8') as f:
            f.write(summary)
            f.write("\n类别列表:\n")
            f.write("-"*60 + "\n")
            for i, cat in enumerate(unique_categories.keys(), 1):
                f.write(f"{i}. {cat}\n")
        
        print_log(f"统计报告已保存: {report_path}")
        
    except Exception as e:
        print_log(f"保存文件失败: {str(e)}")
        return False

    # ===========================================================
    # 分组文本相似度归一化处理
    # ===========================================================
    print_log("\n" + "="*60)
    print_log("开始分组文本相似度归一化处理...")
    print_log("="*60)

    # 1. 获取分组归一化映射
    grouped_maps = group_normalize(df, col_group=group_column_index, col_target=column_index)

    # 2. 准备归一化比对表
    all_mappings = []
    for group_key, mapping in grouped_maps.items():
        for orig, norm in mapping.items():
            # 计算相似度（只基于文字内容）
            orig_text = str(orig)
            norm_text = str(norm)
            if orig_text == norm_text:
                similarity = 100
            else:
                # 使用预处理后的文本计算相似度
                similarity = fuzz.token_sort_ratio(
                    preprocess_text(orig_text),
                    preprocess_text(norm_text)
                )
            all_mappings.append({
                '分组': group_key,
                '原词': orig_text,
                '归一化词': norm_text,
                '相似度(%)': similarity
            })

    # 创建归一化比对表
    normalization_df = pd.DataFrame(all_mappings)

    # 3. 应用归一化到原始数据
    print_log("\n应用归一化到原始数据...")
    start_time = time.time()

    # 创建新列用于存储归一化结果
    df['归一化结果'] = df.iloc[:, column_index]
    normalized_count = 0

    # 使用向量化操作应用归一化
    def apply_normalization(row):
        nonlocal normalized_count
        group_key = row.iloc[group_column_index]
        orig_text = row.iloc[column_index]
        
        # 跳过空值和没有映射的组
        if pd.isna(orig_text) or group_key not in grouped_maps:
            return orig_text
        
        group_map = grouped_maps[group_key]
        normalized_text = group_map.get(orig_text, orig_text)
        
        # 统计变化
        if normalized_text != orig_text:
            normalized_count += 1
            # if normalized_count % 500 == 0:
                # print_log(f"已处理 {normalized_count} 处归一化...")
        
        return normalized_text

    # 应用归一化
    df['归一化结果'] = df.apply(apply_normalization, axis=1)

    # 4. 更新原始列
    df.iloc[:, column_index] = df['归一化结果']
    df = df.drop(columns=['归一化结果'])

    print_log(f"归一化应用完成! 共 {normalized_count} 处修改 | 耗时: {time.time()-start_time:.1f}秒")

    # ===========================================================
    # 保存结果文件
    # ===========================================================
    print_log("\n" + "="*60)
    print_log("保存结果文件...")
    print_log("="*60)

    # 1. 保存归一化比对表（含分组信息）
    normalization_file = os.path.join(output_folder, '分组归一化比对.xlsx')
    try:
        writer = pd.ExcelWriter(normalization_file, engine='xlsxwriter')

        # 按相似度排序
        normalization_df_sorted = normalization_df.sort_values(by='相似度(%)', ascending=False)

        # 将结果写入Excel
        normalization_df_sorted.to_excel(writer, sheet_name='分组归一化映射', index=False)

        # 获取workbook和worksheet对象
        workbook = writer.book
        worksheet = writer.sheets['分组归一化映射']

        # 设置列宽
        worksheet.set_column('A:A', 20)  # 分组列
        worksheet.set_column('B:C', 40)  # 原词和归一化词列
        worksheet.set_column('D:D', 12)  # 相似度列

        # 添加颜色标记
        yellow_format = workbook.add_format({'bg_color': '#FFFF00'})
        normal_format = workbook.add_format()
        blue_format = workbook.add_format({'bg_color': '#CCFFFF'})

        # 添加分组色标
        unique_groups = normalization_df['分组'].unique()
        group_colors = {}
        color_palette = ['#FFCCCC', '#CCFFCC', '#CCCCFF', '#FFFFCC', '#FFCCFF', '#CCFFFF']
        for i, group in enumerate(unique_groups):
            group_colors[group] = workbook.add_format({'bg_color': color_palette[i % len(color_palette)]})

        # 应用格式
        for row_idx in range(1, len(normalization_df_sorted) + 1):
            group_key = normalization_df_sorted.iloc[row_idx - 1]['分组']
            orig_text = normalization_df_sorted.iloc[row_idx - 1]['原词']
            norm_text = normalization_df_sorted.iloc[row_idx - 1]['归一化词']
            
            # 分组列使用颜色
            group_format = group_colors.get(group_key, normal_format)
            worksheet.write(row_idx, 0, group_key, group_format)
            
            # 原词列（如果变化则标黄）
            if orig_text != norm_text:
                worksheet.write(row_idx, 1, orig_text, yellow_format)
                worksheet.write(row_idx, 2, norm_text, yellow_format)
            else:
                worksheet.write(row_idx, 1, orig_text, normal_format)
                worksheet.write(row_idx, 2, norm_text, normal_format)
            
            # 相似度列（数值显示）
            similarity = normalization_df_sorted.iloc[row_idx - 1]['相似度(%)']
            worksheet.write_number(row_idx, 3, similarity)

        # 添加相似度色标
        worksheet.conditional_format(
            f'D2:D{len(normalization_df_sorted)+1}', 
            {
                'type': '2_color_scale',
                'min_value': 0,
                'max_value': 100,
                'min_type': 'num',
                'max_type': 'num',
                'min_color': '#FF0000',  # 红色低相似度
                'max_color': '#00FF00'   # 绿色高相似度
            }
        )

        # 添加标题
        worksheet.write('A1', '分组')
        worksheet.write('B1', '原词')
        worksheet.write('C1', '归一化后的标准词')
        worksheet.write('D1', '相似度 (%)')

        # 添加说明
        worksheet.write('F1', '说明：')
        worksheet.write('F2', '- 黄色标记表示文本被归一化')
        worksheet.write('F3', '- 相似度从红(0%)到绿(100%)渐变')
        worksheet.write('F4', f'- 分组归一化列: {df.columns[group_column_index]}')
        worksheet.write('F5', f'- 阈值设置: 相似度={SIMILARITY_THRESHOLD}% 编辑距离<={EDIT_DISTANCE_THRESHOLD}')
        worksheet.write('F6', f'- 匹配方式: 只匹配文字内容，忽略标点、空格和换行符')

        # 冻结首行
        worksheet.freeze_panes(1, 0)

        writer.close()
        print_log(f"分组归一化比对表已保存到: {normalization_file}")
    except Exception as e:
        print_log(f"保存分组归一化比对表失败: {str(e)}")
        return False

    # 5. 保存最终处理文件
    final_output_path = os.path.join(output_folder, '最终处理_分组归一化.xlsx')
    try:
        final_df = df.copy()
        final_df.to_excel(final_output_path, index=False)
        print_log(f"最终处理文件已保存: {final_output_path}")
    except Exception as e:
        print_log(f"保存最终处理文件失败: {str(e)}")
        return False

    # 最终统计
    print_log("\n" + "="*60)
    print_log("处理结果统计")
    print_log("="*60)
    print_log(f"总行数: {len(df)}")
    
    return True

