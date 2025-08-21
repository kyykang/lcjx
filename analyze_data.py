#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
数据分析脚本 - 分析Excel文件的数据结构
这个脚本用来读取和分析流程效率明细.xls和基本信息.xlsx文件
帮助我们了解数据的结构，为后续的可视化网页开发做准备
"""

import pandas as pd
import sys
import os

def analyze_excel_structure(file_path, file_description):
    """
    分析Excel文件的数据结构
    
    参数:
    file_path: Excel文件的路径
    file_description: 文件描述，用于输出时的标识
    
    功能:
    - 读取Excel文件
    - 显示文件的基本信息（行数、列数、列名等）
    - 显示前几行数据作为样例
    """
    try:
        print(f"\n=== 分析 {file_description} ===")
        print(f"文件路径: {file_path}")
        
        # 尝试不同的方法读取Excel文件
        # 对于.xls文件，可能需要指定engine
        try:
            if file_path.endswith('.xls'):
                # 对于老版本的Excel文件，使用xlrd引擎
                excel_file = pd.ExcelFile(file_path, engine='xlrd')
            else:
                # 对于新版本的Excel文件
                excel_file = pd.ExcelFile(file_path)
        except Exception as e1:
            print(f"尝试使用默认引擎失败: {e1}")
            try:
                # 尝试使用openpyxl引擎
                excel_file = pd.ExcelFile(file_path, engine='openpyxl')
                print("使用openpyxl引擎成功")
            except Exception as e2:
                print(f"尝试使用openpyxl引擎也失败: {e2}")
                return False
        
        print(f"工作表列表: {excel_file.sheet_names}")
        
        # 分析每个工作表
        for sheet_name in excel_file.sheet_names:
            print(f"\n--- 工作表: {sheet_name} ---")
            
            try:
                # 读取工作表数据
                if file_path.endswith('.xls'):
                    df = pd.read_excel(file_path, sheet_name=sheet_name, engine='xlrd')
                else:
                    df = pd.read_excel(file_path, sheet_name=sheet_name)
                
                # 显示基本信息
                print(f"数据形状: {df.shape} (行数: {df.shape[0]}, 列数: {df.shape[1]})")
                print(f"列名: {list(df.columns)}")
                
                # 显示数据类型
                print("\n数据类型:")
                for col in df.columns:
                    print(f"  {col}: {df[col].dtype}")
                
                # 显示前5行数据
                print("\n前5行数据:")
                print(df.head())
                
                # 检查是否有空值
                null_counts = df.isnull().sum()
                if null_counts.sum() > 0:
                    print("\n空值统计:")
                    for col, count in null_counts.items():
                        if count > 0:
                            print(f"  {col}: {count} 个空值")
                else:
                    print("\n没有发现空值")
                    
            except Exception as sheet_error:
                print(f"读取工作表 {sheet_name} 时出错: {str(sheet_error)}")
                continue
                
    except Exception as e:
        print(f"读取文件 {file_path} 时出错: {str(e)}")
        print("可能需要安装xlrd库来读取.xls文件")
        return False
    
    return True

def main():
    """
    主函数 - 分析两个Excel文件的数据结构
    """
    print("开始分析Excel文件数据结构...")
    
    # 设置文件路径
    base_dir = "/Users/kangyiyuan/Desktop/AI编程项目/营销平台流程绩效分析平台"
    
    # 要分析的文件列表
    files_to_analyze = [
        {
            "path": os.path.join(base_dir, "流程效率明细.xls"),
            "description": "流程效率明细文件"
        },
        {
            "path": os.path.join(base_dir, "基本信息.xlsx"),
            "description": "基本信息文件"
        }
    ]
    
    # 分析每个文件
    success_count = 0
    for file_info in files_to_analyze:
        if os.path.exists(file_info["path"]):
            if analyze_excel_structure(file_info["path"], file_info["description"]):
                success_count += 1
        else:
            print(f"文件不存在: {file_info['path']}")
    
    print(f"\n=== 分析完成 ===")
    print(f"成功分析了 {success_count}/{len(files_to_analyze)} 个文件")
    
    if success_count == len(files_to_analyze):
        print("所有文件分析成功！可以开始设计可视化网页了。")
    else:
        print("部分文件分析失败，请检查文件路径和格式。")

if __name__ == "__main__":
    main()