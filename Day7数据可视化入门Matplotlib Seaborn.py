# -*- coding: utf-8 -*-
"""
数据可视化入门（Day 7）
功能：读取销售数据，生成各产品销售金额对比柱状图
包含：解决中文显示问题、数据准备、分组汇总、绘图保存
"""

import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# ==================== 解决中文显示问题 ====================
# 问题：matplotlib 默认字体不支持中文，图表中的中文会显示为方框或乱码
# 解决方法：设置支持中文的字体，如 SimHei（黑体）
plt.rcParams['font.sans-serif'] = ['SimHei']      # 指定默认字体为黑体
plt.rcParams['axes.unicode_minus'] = False        # 解决负号显示异常

# ==================== 1. 读取数据 ====================
# 确保 '销售数据.xlsx' 文件与脚本在同一目录，否则请提供完整路径
df = pd.read_excel('销售数据.xlsx')

# 查看数据基本信息（可选，用于确认列名和数据类型）
print("数据前5行：")
print(df.head())
print("\n数据列信息：")
print(df.info())

# ==================== 2. 数据准备 ====================
# 按产品名称分组，计算总销售额
df_product = df.groupby('产品名称')['金额'].sum().reset_index()
df_product.columns = ['产品名称', '销售额']   # 重命名列，便于绘图

print("\n各产品销售额汇总：")
print(df_product)

# ==================== 3. 绘图 ====================
# 设置画布大小
plt.figure(figsize=(10, 6))

# 使用 seaborn 绘制柱状图
sns.barplot(
    data=df_product,
    x='产品名称',
    y='销售额',
    palette='Blues_d'          # 使用蓝色渐变色系
)

# 添加标题和标签
plt.title('各产品销售金额对比', fontsize=16)
plt.xlabel('产品名称', fontsize=12)
plt.ylabel('销售额（元）', fontsize=12)

# 旋转X轴标签，防止重叠
plt.xticks(rotation=45, ha='right')

# 调整布局，确保标签完整显示
plt.tight_layout()

# ==================== 4. 保存图表 ====================
# 保存为高清PNG图片（dpi=300）
plt.savefig('产品销售金额对比.png', dpi=300, bbox_inches='tight')
print("\n图表已保存为 '产品销售金额对比.png'")

# 显示图表（在脚本运行时弹出窗口）
plt.show()

# ==================== 常见问题及解决方法 ====================
# 1. 如果提示“name 'df' is not defined”：
#    - 检查数据文件是否存在，列名是否正确。
#    - 确认文件路径，建议将数据文件与脚本放在同一目录。
#
# 2. 如果图表中文仍显示为方框：
#    - 可能是系统中没有 SimHei 字体，可尝试其他中文字体，如 'Microsoft YaHei'。
#    - 在代码中替换为 plt.rcParams['font.sans-serif'] = ['Microsoft YaHei']
#
# 3. 如果缺少 seaborn 或 matplotlib 库：
#    - 在终端运行：pip install matplotlib seaborn
#
# 4. 如果保存图片时出现“Permission denied”：
#    - 确保图片文件没有被其他程序打开，或更换保存路径。