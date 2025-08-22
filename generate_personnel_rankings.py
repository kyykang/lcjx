#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
生成人员相关的三个排名数据：
1. 个人流程处理数排名
2. 主要负责人流程数排名  
3. 主要负责人流程处理时长排名
"""

import pandas as pd
import json
from datetime import datetime

def load_personnel_data():
    """加载人员效率明细数据"""
    file_path = '/Users/kangyiyuan/Desktop/AI编程项目/营销平台流程绩效分析平台/人员效率明细.xls'
    
    # 跳过前两行，手动设置列名
    df = pd.read_excel(file_path, header=None, skiprows=2)
    
    # 设置列名
    column_names = [
        '人员名称', '部门名称', '单位名称', '处理数', '平均处理时长', 
        '超期处理数', '超期处理比例', '平均超期时长', '平均超期时长2', 
        '未处理流程数', '超期未处理流程数', '备注'
    ]
    df.columns = column_names[:len(df.columns)]
    
    # 清理数据
    df_clean = df[df['人员名称'].notna() & (df['人员名称'] != '合计')].reset_index(drop=True)
    
    # 转换数值列
    df_clean['处理数_数值'] = pd.to_numeric(df_clean['处理数'], errors='coerce')
    df_clean['未处理流程数_数值'] = pd.to_numeric(df_clean['未处理流程数'], errors='coerce')
    
    return df_clean

def load_main_responsible_persons():
    """加载主要负责人列表"""
    file_path = '/Users/kangyiyuan/Desktop/AI编程项目/营销平台流程绩效分析平台/基本信息.xlsx'
    
    # 读取主要负责人工作表
    df = pd.read_excel(file_path, sheet_name='主要负责人', header=None)
    
    # 主要负责人姓名在第一列，按行排列
    main_persons = df.iloc[:, 0].dropna().tolist()
    
    # 清理姓名，去除可能的空格和特殊字符
    main_persons = [str(name).strip() for name in main_persons if str(name).strip() and str(name) != 'nan']
    
    return main_persons

def generate_personal_process_ranking(df_personnel):
    """生成个人流程处理数排名"""
    # 筛选有效数据（处理数大于0）
    df_valid = df_personnel[df_personnel['处理数_数值'] > 0].copy()
    
    # 按处理数排序
    df_sorted = df_valid.sort_values('处理数_数值', ascending=False)
    
    # 生成排名数据
    ranking_data = []
    for i, (_, row) in enumerate(df_sorted.head(20).iterrows(), 1):
        ranking_data.append({
            '排名': i,
            '人员名称': row['人员名称'],
            '部门名称': row['部门名称'],
            '处理数': int(row['处理数_数值']),
            '平均处理时长': row['平均处理时长'] if pd.notna(row['平均处理时长']) else '-',
            '未处理流程数': int(row['未处理流程数_数值']) if pd.notna(row['未处理流程数_数值']) else 0
        })
    
    return ranking_data

def generate_main_person_process_ranking(df_personnel, main_persons):
    """生成主要负责人流程数排名"""
    # 筛选主要负责人数据
    df_main = df_personnel[df_personnel['人员名称'].isin(main_persons)].copy()
    
    # 筛选有效数据
    df_valid = df_main[df_main['处理数_数值'] > 0].copy()
    
    # 按处理数排序
    df_sorted = df_valid.sort_values('处理数_数值', ascending=False)
    
    # 生成排名数据
    ranking_data = []
    for i, (_, row) in enumerate(df_sorted.head(15).iterrows(), 1):
        ranking_data.append({
            '排名': i,
            '负责人姓名': row['人员名称'],
            '部门名称': row['部门名称'],
            '处理数': int(row['处理数_数值']),
            '平均处理时长': row['平均处理时长'] if pd.notna(row['平均处理时长']) else '-',
            '未处理流程数': int(row['未处理流程数_数值']) if pd.notna(row['未处理流程数_数值']) else 0
        })
    
    return ranking_data

def parse_time_duration(time_str):
    """解析时间字符串，返回分钟数用于排序"""
    if pd.isna(time_str) or time_str == '-' or time_str == '':
        return 0
    
    time_str = str(time_str).strip()
    total_minutes = 0
    
    try:
        # 解析"X天Y小时Z分"格式
        if '天' in time_str:
            parts = time_str.split('天')
            days = int(parts[0])
            total_minutes += days * 24 * 60
            time_str = parts[1] if len(parts) > 1 else ''
        
        if '小时' in time_str:
            parts = time_str.split('小时')
            hours = int(parts[0])
            total_minutes += hours * 60
            time_str = parts[1] if len(parts) > 1 else ''
        
        if '分' in time_str:
            minutes = int(time_str.replace('分', ''))
            total_minutes += minutes
            
    except (ValueError, IndexError):
        return 0
    
    return total_minutes

def generate_main_person_duration_ranking(df_personnel, main_persons):
    """生成主要负责人流程处理时长排名"""
    # 筛选主要负责人数据
    df_main = df_personnel[df_personnel['人员名称'].isin(main_persons)].copy()
    
    # 筛选有处理时长数据的记录
    df_valid = df_main[df_main['平均处理时长'].notna() & (df_main['平均处理时长'] != '-')].copy()
    
    # 解析处理时长为分钟数
    df_valid['处理时长_分钟'] = df_valid['平均处理时长'].apply(parse_time_duration)
    
    # 筛选有效时长数据
    df_valid = df_valid[df_valid['处理时长_分钟'] > 0].copy()
    
    # 按处理时长排序（从长到短）
    df_sorted = df_valid.sort_values('处理时长_分钟', ascending=False)
    
    # 生成排名数据
    ranking_data = []
    for i, (_, row) in enumerate(df_sorted.head(15).iterrows(), 1):
        ranking_data.append({
            '排名': i,
            '负责人姓名': row['人员名称'],
            '部门名称': row['部门名称'],
            '平均处理时长': row['平均处理时长'],
            '处理数': int(row['处理数_数值']) if pd.notna(row['处理数_数值']) else 0,
            '未处理流程数': int(row['未处理流程数_数值']) if pd.notna(row['未处理流程数_数值']) else 0
        })
    
    return ranking_data

def update_chart_data(personal_ranking, main_person_ranking, main_duration_ranking):
    """更新chart_data.json文件"""
    # 读取现有数据
    try:
        with open('chart_data.json', 'r', encoding='utf-8') as f:
            chart_data = json.load(f)
    except FileNotFoundError:
        chart_data = {}
    
    # 添加新的排名数据
    chart_data['personal_process_ranking'] = personal_ranking
    chart_data['main_person_process_ranking'] = main_person_ranking
    chart_data['main_person_duration_ranking'] = main_duration_ranking
    
    # 保存更新后的数据
    with open('chart_data.json', 'w', encoding='utf-8') as f:
        json.dump(chart_data, f, ensure_ascii=False, indent=2)
    
    print("chart_data.json 文件已更新")

def main():
    """主函数"""
    print("开始生成人员排名数据...")
    
    # 加载数据
    print("1. 加载人员效率数据...")
    df_personnel = load_personnel_data()
    print(f"   加载了 {len(df_personnel)} 条人员数据")
    
    print("2. 加载主要负责人列表...")
    main_persons = load_main_responsible_persons()
    print(f"   找到 {len(main_persons)} 位主要负责人")
    print(f"   主要负责人: {main_persons[:10]}...")  # 显示前10个
    
    # 生成排名
    print("3. 生成个人流程处理数排名...")
    personal_ranking = generate_personal_process_ranking(df_personnel)
    print(f"   生成了前 {len(personal_ranking)} 名的排名")
    
    print("4. 生成主要负责人流程数排名...")
    main_person_ranking = generate_main_person_process_ranking(df_personnel, main_persons)
    print(f"   生成了前 {len(main_person_ranking)} 名的排名")
    
    print("5. 生成主要负责人流程处理时长排名...")
    main_duration_ranking = generate_main_person_duration_ranking(df_personnel, main_persons)
    print(f"   生成了前 {len(main_duration_ranking)} 名的排名")
    
    # 更新数据文件
    print("6. 更新chart_data.json文件...")
    update_chart_data(personal_ranking, main_person_ranking, main_duration_ranking)
    
    # 显示示例数据
    print("\n=== 个人流程处理数排名前5名 ===")
    for item in personal_ranking[:5]:
        print(f"{item['排名']}. {item['人员名称']} ({item['部门名称']}) - {item['处理数']}个")
    
    print("\n=== 主要负责人流程数排名前5名 ===")
    for item in main_person_ranking[:5]:
        print(f"{item['排名']}. {item['负责人姓名']} ({item['部门名称']}) - {item['处理数']}个")
    
    print("\n=== 主要负责人流程处理时长排名前5名 ===")
    for item in main_duration_ranking[:5]:
        print(f"{item['排名']}. {item['负责人姓名']} ({item['部门名称']}) - {item['平均处理时长']}")
    
    print("\n数据生成完成！")

if __name__ == "__main__":
    main()