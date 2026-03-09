"""
============================================================
pandas 进阶学习（Day 5） - 财务数据分析常用操作
本脚本包含：读取数据、分组聚合、数据透视表、多表合并、
计算新列、筛选数据、保存结果等。
每个部分都有详细注释，解释代码目的和作用。
使用时请根据实际文件路径和列名进行调整。
============================================================
"""

import pandas as pd

# ==================== 1. 读取数据 ====================
# 目的：从 Excel 文件加载销售数据到 DataFrame
# 假设文件名为 '销售数据_复杂.xlsx'，与脚本在同一目录
# 如果文件在其他位置，请修改为绝对路径，例如 r'D:\data\销售数据.xlsx'
df = pd.read_excel('销售数据_复杂.xlsx')

# 查看数据前5行，快速了解内容
print("=== 数据前5行 ===")
print(df.head())

# 查看数据基本信息：列名、非空数量、数据类型
print("\n=== 数据信息 ===")
df.info()

# 查看数值列的统计摘要（均值、标准差、四分位数等）
print("\n=== 数值列统计摘要 ===")
print(df.describe())

# ==================== 2. 分组聚合（groupby） ====================
# 目的：按某个字段分组，对指定列进行统计计算

# 2.1 单列分组，计算多个统计量
# 按'产品类别'分组，计算'金额'的总和、计数、平均值
category_stats = df.groupby('产品类别')['金额'].agg(['sum', 'count', 'mean'])
print("\n=== 按产品类别分组统计 ===")
print(category_stats)
# 解释：
# - groupby('产品类别')：将相同产品类别的行分为一组
# - ['金额']：指定只对金额列操作
# - agg(['sum','count','mean'])：同时计算总和、计数、平均值

# 2.2 多列分组，并重置索引
# 按'产品类别'和'销售区域'两列分组，计算金额总和
region_category = df.groupby(['产品类别', '销售区域'])['金额'].sum().reset_index()
print("\n=== 按产品类别+销售区域分组统计（已重置索引）===")
print(region_category)
# 解释：
# - groupby(['产品类别','销售区域'])：先按产品类别，再按销售区域分组
# - sum()：计算每组金额总和
# - reset_index()：将分组后作为索引的列恢复为普通列，生成标准DataFrame

# ==================== 3. 数据透视表（pivot_table） ====================
# 目的：类似 Excel 透视表，行列交叉汇总
pivot = pd.pivot_table(
    df,
    values='金额',               # 要汇总的数值列
    index='产品类别',            # 行字段
    columns='销售区域',          # 列字段
    aggfunc='sum',               # 汇总方式：求和
    fill_value=0,                # 缺失值填充为0
    margins=True,                # 添加总计行/列
    margins_name='合计'          # 总计行列的名称
)
print("\n=== 数据透视表（产品类别 vs 销售区域）===")
print(pivot)
# 解释：
# - values：要聚合的列
# - index：放在行位置的列
# - columns：放在列位置的列
# - aggfunc：可选的函数，如'sum','mean','count'等，也可传入自定义函数
# - fill_value：当交叉点无数据时填充的值，避免NaN
# - margins：是否显示行/列总计
# - margins_name：总计的名称

# ==================== 4. 多表合并（merge） ====================
# 目的：将销售表与产品信息表合并，获得成本价、供应商等字段

# 4.1 创建产品信息表（示例数据）
products = pd.DataFrame({
    '产品名称': ['笔记本', '显示器', '键盘', '笔', '文件夹'],
    '成本价': [3000, 1200, 180, 2, 8],
    '供应商': ['宏达科技', '华星光电', '达尔优', '晨光', '得力']
})
print("\n=== 产品信息表 ===")
print(products)

# 4.2 左连接合并（保留销售表所有行，产品信息匹配不上则显示NaN）
df_merged = pd.merge(
    df,
    products,
    on='产品名称',      # 连接键，两表该列名称必须相同
    how='left'          # 连接方式：left左连接，保留左表所有行
)
print("\n=== 合并后的数据（前5行）===")
print(df_merged.head())
# 解释：
# - on：指定连接字段，两表必须都有此列
# - how：连接方式
#   * 'inner'：只保留两表都有的记录
#   * 'left'：保留左表全部，右表匹配不上为NaN
#   * 'right'：保留右表全部
#   * 'outer'：保留两表全部

# ==================== 5. 计算新列 ====================
# 目的：基于已有列生成新的指标列

# 计算毛利 = 金额 - 成本价 × 销售数量
df_merged['毛利'] = df_merged['金额'] - df_merged['成本价'] * df_merged['销售数量']

# 计算毛利率 = 毛利 / 金额（注意除零问题，可能产生inf）
df_merged['毛利率'] = df_merged['毛利'] / df_merged['金额']

print("\n=== 添加毛利和毛利率后的数据（前5行）===")
print(df_merged[['产品名称', '金额', '成本价', '销售数量', '毛利', '毛利率']].head())
# 解释：
# - 直接通过列名赋值创建新列
# - 运算基于向量化操作，一行代码对所有行生效
# - 如果金额为0，毛利率可能为无穷大，后续可用 fillna 处理

# ==================== 6. 筛选数据 ====================
# 目的：根据条件提取特定行

# 6.1 简单筛选：金额 > 5000
df_big = df[df['金额'] > 5000]
print("\n=== 金额 > 5000 的记录 ===")
print(df_big.head())

# 6.2 多条件筛选：电子产品 且 金额 > 5000
df_electronics_big = df[(df['产品类别'] == '电子产品') & (df['金额'] > 5000)]
print("\n=== 电子产品且金额 > 5000 的记录 ===")
print(df_electronics_big.head())

# 6.3 使用 isin() 筛选多个值
df_keyboard = df[df['产品名称'].isin(['键盘', '显示器'])]
print("\n=== 产品名称为键盘或显示器的记录 ===")
print(df_keyboard.head())

# 6.4 使用 query() 方法（类似SQL风格）
df_query = df.query('金额 > 5000 and 产品类别 == "电子产品"')
print("\n=== 使用 query 筛选的结果 ===")
print(df_query.head())
# 解释：
# - df[条件] 返回满足条件的行，条件必须写在方括号内
# - 多条件用 &（与）、|（或），每个条件必须用括号括起
# - isin() 检查某列是否在给定列表内
# - query() 直接用字符串写条件，列名不需加引号，字符串值需加引号

# ==================== 7. 保存结果 ====================
# 目的：将处理后的数据或分析结果保存为 Excel 文件

# 7.1 保存透视表
pivot.to_excel('销售透视表_pandas.xlsx', sheet_name='透视表结果')
print("\n=== 透视表已保存为 '销售透视表_pandas.xlsx' ===")

# 7.2 保存合并后的数据（不保存行索引）
df_merged.to_excel('销售与成本合并表.xlsx', index=False)
print("合并表已保存为 '销售与成本合并表.xlsx'")

# 7.3 同时保存多个 DataFrame 到同一个 Excel 的不同工作表
with pd.ExcelWriter('pandas_练习结果.xlsx') as writer:
    category_stats.to_excel(writer, sheet_name='产品类别统计')
    pivot.to_excel(writer, sheet_name='透视表')
    df_merged.to_excel(writer, sheet_name='合并数据', index=False)
print("多工作表文件已保存为 'pandas_练习结果.xlsx'")

# ==================== 8. 完整练习任务 ====================
# 目的：综合运用所学知识完成一个财务分析小任务

# 8.1 按销售员分组，计算总销售额和订单数
salesman_stats = df.groupby('销售员')['金额'].agg(['sum', 'count']).reset_index()
print("\n=== 销售员业绩统计 ===")
print(salesman_stats)

# 8.2 用透视表展示各销售员在各产品类别的销售额
pivot_salesman = pd.pivot_table(
    df,
    values='金额',
    index='销售员',
    columns='产品类别',
    aggfunc='sum',
    fill_value=0
)
print("\n=== 销售员 vs 产品类别销售额透视表 ===")
print(pivot_salesman)

# 8.3 将两个结果保存到同一 Excel 文件的不同工作表
with pd.ExcelWriter('销售员分析结果.xlsx') as writer:
    salesman_stats.to_excel(writer, sheet_name='销售员业绩', index=False)
    pivot_salesman.to_excel(writer, sheet_name='销售员产品透视')
print("销售员分析结果已保存为 '销售员分析结果.xlsx'")

print("\n=== 所有任务完成 ===")