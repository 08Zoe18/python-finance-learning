import pandas as pd

# 读取 Excel
df = pd.read_excel('销售数据.xlsx')
print(df.head())          # 查看前5行
print(df.info())          # 查看列名、数据类型、非空计数
print(df.describe())      # 数值列的统计摘要
# 筛选出金额大于5000的行
df_big = df[df['金额'] > 5000]
print(df_big)

# 多条件筛选：金额>5000 且 销售区域='华东'
df_huadong_big = df[(df['金额'] > 5000) & (df['销售区域'] == '华东')]
print(df_huadong_big)
# 按产品类别分组，计算总销售额
category_sum = df.groupby('产品类别')['金额'].sum()
print(category_sum)

# 按产品类别和销售区域分组，计算总销售额
region_category = df.groupby(['产品类别', '销售区域'])['金额'].sum()
print(region_category)

# 重置索引，变回普通 DataFrame
region_category_df = region_category.reset_index()

# 透视表：行=产品类别，列=销售区域，值=金额（求和）
pivot = pd.pivot_table(df, 
                       values='金额', 
                       index='产品类别', 
                       columns='销售区域', 
                       aggfunc='sum', 
                       fill_value=0)
print(pivot)

pivot.to_excel('销售透视表_pandas.xlsx')
print('已保存')