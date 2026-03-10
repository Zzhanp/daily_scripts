#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
供应商物料分析脚本
功能：
1. 根据四分位法将生产发运时长分为三个阶段
2. 计算各阶段的平均值和最小值
3. 计算各阶段的最小需求值
4. 如果阶段一和阶段二最小需求值相同，则阶段一需求设为0
"""

import pandas as pd
import numpy as np

# 读取Excel
input_file = '/home/node/.openclaw/media/inbound/生产-1---436b92d0-11d7-47a0-a8b7-04a33eb9ccb0.xlsx'
output_file = '/home/node/.openclaw/workspace/供应商物料分析结果.xlsx'

df = pd.read_excel(input_file)

# 清理列名（去除空格）
df.columns = df.columns.str.strip()

print("=" * 60)
print("原始数据概览")
print("=" * 60)
print(f"总记录数: {len(df)}")
print(f"列名: {df.columns.tolist()}")

# 按供应商和物料分组分析
results = []

# 获取所有供应商-物料组合
groups = df.groupby(['供应商信息', '物资品类'])

for (supplier, material), group in groups:
    print(f"\n{'='*60}")
    print(f"供应商: {supplier}, 物料: {material}")
    print(f"记录数: {len(group)}")
    
    # 按生产发运时长排序
    group_sorted = group.sort_values('生产发运时长').reset_index(drop=True)
    
    # 计算四分位数
    durations = group_sorted['生产发运时长'].values
    demands = group_sorted['求和项:需求数量'].values
    
    Q1 = np.percentile(durations, 25)
    Q2 = np.percentile(durations, 50)  # 中位数
    Q3 = np.percentile(durations, 75)
    
    print(f"四分位数: Q1={Q1:.2f}, Q2={Q2:.2f}, Q3={Q3:.2f}")
    
    # 分成三个阶段
    stage1_idx = durations < Q1
    stage2_idx = (durations >= Q1) & (durations < Q2)
    stage3_idx = durations >= Q2
    
    # 阶段1分析
    stage1_durations = durations[stage1_idx]
    stage1_demands = demands[stage1_idx]
    stage1_avg_duration = np.mean(stage1_durations) if len(stage1_durations) > 0 else 0
    stage1_min_duration = np.min(stage1_durations) if len(stage1_durations) > 0 else 0
    stage1_min_demand = np.min(stage1_demands) if len(stage1_demands) > 0 else 0
    
    # 阶段2分析
    stage2_durations = durations[stage2_idx]
    stage2_demands = demands[stage2_idx]
    stage2_avg_duration = np.mean(stage2_durations) if len(stage2_durations) > 0 else 0
    stage2_min_duration = np.min(stage2_durations) if len(stage2_durations) > 0 else 0
    stage2_min_demand = np.min(stage2_demands) if len(stage2_demands) > 0 else 0
    
    # 阶段3分析
    stage3_durations = durations[stage3_idx]
    stage3_demands = demands[stage3_idx]
    stage3_avg_duration = np.mean(stage3_durations) if len(stage3_durations) > 0 else 0
    stage3_min_duration = np.min(stage3_durations) if len(stage3_durations) > 0 else 0
    stage3_min_demand = np.min(stage3_demands) if len(stage3_demands) > 0 else 0
    
    print(f"\n阶段1 (时长 < {Q1:.2f}): 记录数={len(stage1_durations)}")
    print(f"  平均时长: {stage1_avg_duration:.2f}, 最小时长: {stage1_min_duration:.2f}")
    print(f"  最小需求: {stage1_min_demand}")
    
    print(f"\n阶段2 (时长 {Q1:.2f} - {Q2:.2f}): 记录数={len(stage2_durations)}")
    print(f"  平均时长: {stage2_avg_duration:.2f}, 最小时长: {stage2_min_duration:.2f}")
    print(f"  最小需求: {stage2_min_demand}")
    
    print(f"\n阶段3 (时长 >= {Q2:.2f}): 记录数={len(stage3_durations)}")
    print(f"  平均时长: {stage3_avg_duration:.2f}, 最小时长: {stage3_min_duration:.2f}")
    print(f"  最小需求: {stage3_min_demand}")
    
    # 规则4: 如果阶段一和阶段二最小需求值相同，则阶段一需求为0
    if stage1_min_demand == stage2_min_demand:
        print(f"\n*** 规则应用: 阶段1最小需求 = 阶段2最小需求 = {stage1_min_demand}, 将阶段1设为0 ***")
        stage1_min_demand_final = 0
    else:
        stage1_min_demand_final = stage1_min_demand
    
    # 保存结果
    results.append({
        '供应商ID': supplier,
        '供应商名称': material.split()[0] if len(material.split()) > 1 else '',
        '物料': material,
        '阶段': '阶段1',
        '记录数': len(stage1_durations),
        '平均时长': round(stage1_avg_duration, 2),
        '最小时长': round(stage1_min_duration, 2),
        '最小需求': stage1_min_demand_final
    })
    
    results.append({
        '供应商ID': supplier,
        '供应商名称': material.split()[0] if len(material.split()) > 1 else '',
        '物料': material,
        '阶段': '阶段2',
        '记录数': len(stage2_durations),
        '平均时长': round(stage2_avg_duration, 2),
        '最小时长': round(stage2_min_duration, 2),
        '最小需求': stage2_min_demand
    })
    
    results.append({
        '供应商ID': supplier,
        '供应商名称': material.split()[0] if len(material.split()) > 1 else '',
        '物料': material,
        '阶段': '阶段3',
        '记录数': len(stage3_durations),
        '平均时长': round(stage3_avg_duration, 2),
        '最小时长': round(stage3_min_duration, 2),
        '最小需求': stage3_min_demand
    })

# 创建结果DataFrame
result_df = pd.DataFrame(results)

print("\n" + "=" * 60)
print("最终结果")
print("=" * 60)
print(result_df.to_string(index=False))

# 输出到Excel
result_df.to_excel(output_file, index=False, sheet_name='分析结果')
print(f"\n结果已保存到: {output_file}")
