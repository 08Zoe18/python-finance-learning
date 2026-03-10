"""
pandas 数据处理进阶 - Day 6
包含：缺失值处理、重复值处理、数据类型转换、异常值处理、标准化、字符串操作
"""
import pandas as pd

# ========== 1. 读取数据 ==========
df = pd.read_excel('销售数据_复杂.xlsx')
print("原始数据形状：", df.shape)

# ========== 2. 缺失值处理 ==========
print("缺失值统计：\n", df.isnull().sum())

# 方法1：删除缺失值
df_dropna = df.dropna()
print("删除缺失值后形状：", df_dropna.shape)

# 方法2：填充缺失值（用均值填充数值列）
for col in df.select_dtypes(include=['float64', 'int64']).columns:
    df[col].fillna(df[col].mean(), inplace=True)

# 方法3：向前填充（时间序列常用）
df.sort_values('日期', inplace=True)
df.fillna(method='ffill', inplace=True)

# ========== 3. 重复值处理 ==========
print("重复行数：", df.duplicated().sum())
df.drop_duplicates(inplace=True)

# ========== 4. 数据类型转换 ==========
print("数据类型：\n", df.dtypes)
df['日期'] = pd.to_datetime(df['日期'])
df['金额'] = pd.to_numeric(df['金额'], errors='coerce')
df['产品类别'] = df['产品类别'].astype('category')

# ========== 5. 异常值处理（以金额为例） ==========
Q1 = df['金额'].quantile(0.25)
Q3 = df['金额'].quantile(0.75)
IQR = Q3 - Q1
lower = Q1 - 1.5 * IQR
upper = Q3 + 1.5 * IQR
outliers = df[(df['金额'] < lower) | (df['金额'] > upper)]
print(f"异常值数量：{len(outliers)}")
# 截尾处理
df['金额'] = df['金额'].clip(lower, upper)

# ========== 6. 数据标准化 ==========
df['金额_标准化'] = (df['金额'] - df['金额'].min()) / (df['金额'].max() - df['金额'].min())
df['金额_zscore'] = (df['金额'] - df['金额'].mean()) / df['金额'].std()

# ========== 7. 字符串操作 ==========
df['产品名称'] = df['产品名称'].str.strip().str.title()

# ========== 8. 保存结果 ==========
with pd.ExcelWriter('清洗后数据.xlsx') as writer:
    df.to_excel(writer, sheet_name='清洗后', index=False)
    df.describe().to_excel(writer, sheet_name='统计描述')
print("数据清洗完成，已保存。")