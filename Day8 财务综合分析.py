# -*- coding: utf-8 -*-
"""
金蝶报表综合财务分析
功能：
- 读取利润表、科目余额表
- 提取月度收入、成本、费用、利润
- 生成趋势图、结构图、对比图
- 保存图表到本地
"""

import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import os

# 设置中文字体
plt.rcParams['font.sans-serif'] = ['SimHei']
plt.rcParams['axes.unicode_minus'] = False

# 创建保存图表的文件夹
os.makedirs('财务分析图表', exist_ok=True)

# ==================== 1. 读取利润表 ====================
# 利润表通常有“项目”列和“本年累计”、“本月金额”等列
try:
    profit_df = pd.read_excel('利润表_2025全年.xlsx')
    print("利润表前5行：")
    print(profit_df.head())
except FileNotFoundError:
    print("未找到利润表文件，请检查路径。")
    # 如果没有真实数据，可以用示例数据代替
    profit_df = pd.DataFrame({
        '项目': ['营业收入', '营业成本', '管理费用', '财务费用', '营业利润'],
        '1月': [100000, 60000, 15000, 500, 24500],
        '2月': [120000, 72000, 16000, 600, 31400],
        '3月': [130000, 78000, 15500, 550, 35950],
    })
    print("使用示例数据：")
    print(profit_df)

# ==================== 2. 读取科目余额表 ====================
# 科目余额表通常有“科目名称”、“期初余额”、“本期借方”、“本期贷方”、“期末余额”等
try:
    balance_df = pd.read_excel('科目余额表_2025全年.xlsx')
    print("科目余额表前5行：")
    print(balance_df.head())
except FileNotFoundError:
    balance_df = None
    print("未找到科目余额表，跳过部分分析。")

# ==================== 3. 数据整理 ====================
# 将利润表转为月度格式（假设列为月份）
# 利润表通常是项目为行，月份为列
if '项目' in profit_df.columns:
    # 转置，以月份为行
    profit_t = profit_df.set_index('项目').T
    profit_t.index.name = '月份'
    profit_t.reset_index(inplace=True)
    # 将月份转为字符串，便于绘图
    profit_t['月份'] = profit_t['月份'].astype(str)
else:
    profit_t = profit_df  # 如果已经是月度格式

print("整理后的利润表：")
print(profit_t.head())

# ==================== 4. 计算关键指标 ====================
# 假设利润表中有“营业收入”、“营业成本”、“管理费用”、“财务费用”、“营业利润”
# 如果列名不一致，请根据实际修改
revenue_col = '营业收入' if '营业收入' in profit_t.columns else None
cost_col = '营业成本' if '营业成本' in profit_t.columns else None
expense_cols = ['管理费用', '财务费用'] if '管理费用' in profit_t.columns else []
profit_col = '营业利润' if '营业利润' in profit_t.columns else None

# 计算总费用
if expense_cols:
    profit_t['总费用'] = profit_t[expense_cols].sum(axis=1)

# ==================== 5. 绘制收入趋势图 ====================
if revenue_col:
    plt.figure(figsize=(10, 5))
    plt.plot(profit_t['月份'], profit_t[revenue_col], marker='o', linestyle='-', color='b', label='营业收入')
    plt.title('2025年各月收入趋势')
    plt.xlabel('月份')
    plt.ylabel('金额（元）')
    plt.grid(True, linestyle='--', alpha=0.6)
    plt.legend()
    plt.tight_layout()
    plt.savefig('财务分析图表/月度收入趋势.png', dpi=300)
    plt.show()

# ==================== 6. 绘制费用结构饼图 ====================
if expense_cols and len(expense_cols) > 0:
    # 计算全年各项费用总和
    expense_sum = profit_t[expense_cols].sum()
    plt.figure(figsize=(8, 8))
    plt.pie(expense_sum, labels=expense_sum.index, autopct='%1.1f%%', startangle=90)
    plt.title('2025年费用结构')
    plt.tight_layout()
    plt.savefig('财务分析图表/费用结构饼图.png', dpi=300)
    plt.show()

# ==================== 7. 绘制月度利润柱状图 ====================
if profit_col:
    plt.figure(figsize=(12, 6))
    sns.barplot(data=profit_t, x='月份', y=profit_col, palette='viridis')
    plt.title('2025年各月营业利润')
    plt.xlabel('月份')
    plt.ylabel('利润（元）')
    plt.xticks(rotation=45)
    plt.tight_layout()
    plt.savefig('财务分析图表/月度利润柱状图.png', dpi=300)
    plt.show()

# ==================== 8. 如果有产品数据，可以绘制产品毛利对比 ====================
if balance_df is not None:
    # 从科目余额表提取主营业务收入和成本明细（假设有产品明细科目）
    # 这里需要根据你的实际科目名称调整
    income_products = balance_df[balance_df['科目名称'].str.contains('主营业务收入', na=False)]
    cost_products = balance_df[balance_df['科目名称'].str.contains('主营业务成本', na=False)]

    if not income_products.empty and not cost_products.empty:
        # 提取产品名称（假设科目名称格式为“主营业务收入—A产品”）
        income_products['产品'] = income_products['科目名称'].str.extract(r'—(.+)')
        cost_products['产品'] = cost_products['科目名称'].str.extract(r'—(.+)')

        # 合并收入和成本
        merged = pd.merge(income_products[['产品', '本年累计']], cost_products[['产品', '本年累计']],
                          on='产品', how='outer', suffixes=('_收入', '_成本')).fillna(0)
        merged['毛利'] = merged['本年累计_收入'] - merged['本年累计_成本']
        merged['毛利率'] = merged['毛利'] / merged['本年累计_收入'] * 100

        # 绘制毛利对比柱状图
        plt.figure(figsize=(10, 6))
        x = range(len(merged))
        plt.bar(x, merged['毛利'], color='skyblue', label='毛利')
        plt.xticks(x, merged['产品'], rotation=45)
        plt.title('各产品毛利对比')
        plt.ylabel('毛利（元）')
        for i, v in enumerate(merged['毛利']):
            plt.text(i, v + 100, f'{v:,.0f}', ha='center', va='bottom')
        plt.tight_layout()
        plt.savefig('财务分析图表/产品毛利对比.png', dpi=300)
        plt.show()

print("所有图表已生成，保存在 '财务分析图表' 文件夹中。")