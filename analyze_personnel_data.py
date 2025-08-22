#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
分析人员效率明细.xls文件的数据结构
"""

import pandas as pd

def analyze_personnel_data():
    """分析人员效率明细Excel文件"""
    file_path = '/Users/kangyiyuan/Desktop/AI编程项目/营销平台流程绩效分析平台/人员效率明细.xls'
    
    # 先读取原始数据看结构
    print("=== 原始数据结构 ===")
    df_raw = pd.read_excel(file_path)
    print(f"原始数据形状: {df_raw.shape}")
    print("前10行:")
    print(df_raw.head(10))
    
    # 手动设置列名，基于观察到的数据结构
    print("\n=== 手动设置列名 ===")
    df = pd.read_excel(file_path, header=None, skiprows=2)  # 跳过前两行
    
    # 根据观察到的数据结构设置列名
    column_names = [
        '人员名称', '部门名称', '单位名称', '处理数', '平均处理时长', 
        '超期处理数', '超期处理比例', '平均超期时长', '未处理流程数', '超期未处理流程数'
    ]
    
    # 如果列数不匹配，调整列名
    if len(df.columns) != len(column_names):
        print(f"实际列数: {len(df.columns)}, 预期列数: {len(column_names)}")
        # 根据实际列数调整
        if len(df.columns) == 12:
            column_names = [
                '人员名称', '部门名称', '单位名称', '处理数', '平均处理时长', 
                '超期处理数', '超期处理比例', '平均超期时长', '平均超期时长2', 
                '未处理流程数', '超期未处理流程数', '备注'
            ]
    
    df.columns = column_names[:len(df.columns)]
    
    print(f"数据形状: {df.shape}")
    print("列名:")
    for i, col in enumerate(df.columns):
        print(f"{i}: '{col}'")
    
    print("\n前5行数据:")
    print(df.head())
    
    # 清理数据，去除合计行和空行
    df_clean = df[df['人员名称'].notna() & (df['人员名称'] != '合计')].reset_index(drop=True)
    print(f"\n去除合计行后数据形状: {df_clean.shape}")
    print("清理后前10行:")
    print(df_clean.head(10))
    
    # 检查处理数列的数据
    print("\n处理数列的唯一值示例:")
    print(df_clean['处理数'].unique()[:20])
    
    # 转换处理数为数值类型
    df_clean['处理数_数值'] = pd.to_numeric(df_clean['处理数'], errors='coerce')
    
    # 去除处理数为0或空值的记录
    df_valid = df_clean[df_clean['处理数_数值'] > 0].copy()
    
    print(f"\n有效数据行数: {len(df_valid)}")
    print("处理数统计:")
    print(df_valid['处理数_数值'].describe())
    
    # 按处理数排序看排名
    print("\n个人流程处理数排名前10:")
    df_sorted = df_valid.sort_values('处理数_数值', ascending=False)
    print(df_sorted[['人员名称', '部门名称', '处理数', '处理数_数值']].head(10))
    
    return df_clean

if __name__ == "__main__":
    analyze_personnel_data()