#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
分析基本信息.xlsx文件，获取主要负责人信息
"""

import pandas as pd

def analyze_basic_info():
    """分析基本信息Excel文件"""
    file_path = '/Users/kangyiyuan/Desktop/AI编程项目/营销平台流程绩效分析平台/基本信息.xlsx'
    
    try:
        # 检查Excel文件的所有工作表
        print("=== 检查Excel工作表 ===")
        excel_file = pd.ExcelFile(file_path)
        print(f"工作表名称: {excel_file.sheet_names}")
        
        # 专门分析主要负责人工作表
        if '主要负责人' in excel_file.sheet_names:
            print("\n=== 详细分析主要负责人工作表 ===")
            
            # 尝试不同的读取方式
            df_raw = pd.read_excel(file_path, sheet_name='主要负责人', header=None)
            print(f"原始数据形状: {df_raw.shape}")
            print("前10行原始数据:")
            print(df_raw.head(10))
            
            # 尝试转置数据（可能是横向排列的）
            df_transposed = df_raw.T
            print(f"\n转置后数据形状: {df_transposed.shape}")
            print("转置后前10行:")
            print(df_transposed.head(10))
            
            # 检查是否有多列数据
            print("\n各列非空值数量:")
            for i, col in enumerate(df_raw.columns):
                non_null_count = df_raw[col].notna().sum()
                print(f"列{i}: {non_null_count}个非空值")
                if non_null_count > 0:
                    print(f"  前5个非空值: {df_raw[col].dropna().head().tolist()}")
        
        # 读取其他工作表看是否有流程信息
        for sheet_name in excel_file.sheet_names:
            if sheet_name != '主要负责人':
                print(f"\n=== 工作表: {sheet_name} ===")
                df = pd.read_excel(file_path, sheet_name=sheet_name)
                print(f"数据形状: {df.shape}")
                print("流程列表:")
                if len(df.columns) > 0:
                    process_list = df.iloc[:, 0].dropna().tolist()
                    print(process_list[:10])  # 显示前10个流程
        
        # 尝试不同的读取方式
        print("\n=== 尝试不同的读取方式 ===")
        # 尝试读取所有列
        df_all = pd.read_excel(file_path, header=None)
        print(f"无表头读取数据形状: {df_all.shape}")
        print("前10行:")
        print(df_all.head(10))
        
        return excel_file.sheet_names
        
    except Exception as e:
        print(f"读取文件时出错: {e}")
        return None

if __name__ == "__main__":
    analyze_basic_info()