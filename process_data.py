#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
数据处理脚本 - 将Excel数据转换为JSON格式
这个脚本读取Excel文件并生成可视化网页所需的JSON数据
"""

import pandas as pd
import json
import os
from datetime import datetime

def clean_column_names(df):
    """
    清理DataFrame的列名
    将Unnamed列替换为更有意义的名称
    """
    # 根据分析结果，流程效率明细的列应该是：
    expected_columns = [
        '模板名称', '发起流程数', '环比', '同比', '完成流程数', 
        '完成环比', '平均运行时长', '运行时长环比', '超期结束比例', 
        '平均超期时长', '未结束流程数', '超期未结束流程数', '备注'
    ]
    
    # 如果列数匹配，就使用预期的列名
    if len(df.columns) == len(expected_columns):
        df.columns = expected_columns
    
    return df

def process_flow_efficiency_data():
    """
    处理流程效率明细数据
    返回处理后的数据字典
    """
    file_path = "/Users/kangyiyuan/Desktop/AI编程项目/营销平台流程绩效分析平台/流程效率明细.xls"
    
    try:
        # 读取Excel文件
        df = pd.read_excel(file_path, engine='xlrd')
        
        # 清理列名
        df = clean_column_names(df)
        
        # 删除前两行（标题行）和最后的汇总行
        df = df.iloc[2:].reset_index(drop=True)
        
        # 删除"合计"行
        df = df[df['模板名称'] != '合计'].reset_index(drop=True)
        
        # 清理数据，移除空值行
        df = df.dropna(subset=['模板名称'])
        
        # 转换数值列
        numeric_columns = ['发起流程数', '完成流程数', '未结束流程数', '超期未结束流程数']
        for col in numeric_columns:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
        
        # 处理平均运行时长（可能包含"天"、"小时"等单位）
        if '平均运行时长' in df.columns:
            df['平均运行时长_数值'] = df['平均运行时长'].apply(parse_duration)
        
        # 转换为字典格式
        data = df.to_dict('records')
        
        return {
            'success': True,
            'data': data,
            'total_records': len(data)
        }
        
    except Exception as e:
        return {
            'success': False,
            'error': str(e),
            'data': []
        }

def parse_duration(duration_str):
    """
    解析时长字符串，转换为小时数
    例如："2天3小时" -> 51小时
    """
    if pd.isna(duration_str) or duration_str == '-':
        return 0
    
    try:
        duration_str = str(duration_str)
        hours = 0
        
        # 提取天数
        if '天' in duration_str:
            days_part = duration_str.split('天')[0]
            if days_part.isdigit():
                hours += int(days_part) * 24
        
        # 提取小时数
        if '小时' in duration_str:
            if '天' in duration_str:
                hours_part = duration_str.split('天')[1].split('小时')[0]
            else:
                hours_part = duration_str.split('小时')[0]
            
            if hours_part.isdigit():
                hours += int(hours_part)
        
        # 如果只是数字，假设是小时
        elif duration_str.replace('.', '').isdigit():
            hours = float(duration_str)
        
        return hours
        
    except:
        return 0

def process_flow_categories():
    """
    处理流程分类数据
    返回各类流程的列表
    """
    file_path = "/Users/kangyiyuan/Desktop/AI编程项目/营销平台流程绩效分析平台/基本信息.xlsx"
    
    try:
        categories = {}
        
        # 读取各个工作表
        sheet_names = ['销售类流程', '采购类流程', '项目&产品管理类流程']
        
        for sheet_name in sheet_names:
            df = pd.read_excel(file_path, sheet_name=sheet_name)
            # 获取第一列的所有非空值
            flow_list = df.iloc[:, 0].dropna().tolist()
            categories[sheet_name] = flow_list
        
        return {
            'success': True,
            'categories': categories
        }
        
    except Exception as e:
        return {
            'success': False,
            'error': str(e),
            'categories': {}
        }

def generate_chart_data():
    """
    生成图表所需的数据
    """
    # 获取流程效率数据
    efficiency_result = process_flow_efficiency_data()
    categories_result = process_flow_categories()
    
    if not efficiency_result['success'] or not categories_result['success']:
        return {
            'success': False,
            'error': 'Failed to process data'
        }
    
    data = efficiency_result['data']
    categories = categories_result['categories']
    
    # 1. 发起流程数及完成数排名
    flow_ranking = sorted(data, key=lambda x: x.get('发起流程数', 0), reverse=True)[:10]
    
    # 2. 流程运行时长排名
    duration_ranking = sorted(
        [item for item in data if item.get('平均运行时长_数值', 0) > 0], 
        key=lambda x: x.get('平均运行时长_数值', 0), 
        reverse=True
    )[:10]
    
    # 3-5. 各类流程的平均运行时长排名
    category_rankings = {}
    for category_name, flow_list in categories.items():
        category_data = [item for item in data if item.get('模板名称') in flow_list]
        category_ranking = sorted(
            [item for item in category_data if item.get('平均运行时长_数值', 0) > 0],
            key=lambda x: x.get('平均运行时长_数值', 0),
            reverse=True
        )[:10]
        category_rankings[category_name] = category_ranking
    
    return {
        'success': True,
        'data': {
            'flow_ranking': flow_ranking,
            'duration_ranking': duration_ranking,
            'category_rankings': category_rankings,
            'categories': categories,
            'raw_data': data
        },
        'generated_at': datetime.now().isoformat()
    }

def main():
    """
    主函数 - 生成JSON数据文件
    """
    print("开始处理数据...")
    
    # 生成图表数据
    result = generate_chart_data()
    
    if result['success']:
        # 保存为JSON文件
        output_file = "/Users/kangyiyuan/Desktop/AI编程项目/营销平台流程绩效分析平台/chart_data.json"
        
        with open(output_file, 'w', encoding='utf-8') as f:
            json.dump(result, f, ensure_ascii=False, indent=2)
        
        print(f"数据处理完成！")
        print(f"输出文件: {output_file}")
        print(f"发起流程数排名前10: {len(result['data']['flow_ranking'])} 条记录")
        print(f"运行时长排名前10: {len(result['data']['duration_ranking'])} 条记录")
        
        for category, ranking in result['data']['category_rankings'].items():
            print(f"{category}排名: {len(ranking)} 条记录")
            
    else:
        print(f"数据处理失败: {result.get('error', '未知错误')}")

if __name__ == "__main__":
    main()