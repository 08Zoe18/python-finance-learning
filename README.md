# python-finance-learning
## Day 1（2026.03.05）
- 安装Python环境
- 编写并运行第一个程序 Hello.py
## Day 2（2026.03.06）
- 学习内容：用Pandas读取金蝶明细账，进行产品收入成本分析
- 代码存档：`/Day2/analyze_3month_fixed.py`
## Day 3（2026.03.07）
- 编写了实用脚本 `data_clean.py`：自动删除重复行 + 按产品汇总金额
- 文件位置：`/Day3_实用脚本/data_clean.py`
## Day 4（2026.03.08）
- 学习内容：pandas 基础（读取、筛选、分组、透视表）
- 练习文件：[pandas_day4.py](./Day4/pandas_day4.py)
- 生成结果：[销售透视表_pandas.xlsx](./Day4/销售透视表_pandas.xlsx)
- 收获：能用 pandas 完成数据筛选、分组汇总和透视表，比 Excel 更高效。
## Day 5（2026.03.09）
- **学习内容**：pandas 进阶操作（分组聚合、数据透视表、多表合并、计算新列、筛选数据）
- **练习文件**：[pandas_day5.py](./Day5/pandas_day5.py)
- **核心收获**：
  - 掌握 `groupby` 的单列/多列分组及多种聚合函数（sum/count/mean）
  - 学会用 `pivot_table` 创建数据透视表，添加总计行列
  - 能够用 `merge` 实现类似 SQL 的表连接（inner/left/right/outer）
  - 熟练计算新列（毛利、毛利率）并进行条件筛选
  - 能将多个结果保存到同一 Excel 的不同工作表
- **示例代码概览**：
  ```python
  # 分组聚合
  category_stats = df.groupby('产品类别')['金额'].agg(['sum','count','mean'])

  # 数据透视表
  pivot = pd.pivot_table(df, values='金额', index='产品类别', columns='销售区域', aggfunc='sum', fill_value=0, margins=True)

  # 多表合并
  df_merged = pd.merge(df, products, on='产品名称', how='left')

  # 计算新列
  df_merged['毛利'] = df_merged['金额'] - df_merged['成本价'] * df_merged['销售数量']
  df_merged['毛利率'] = df_merged['毛利'] / df_merged['金额']

  # 保存多工作表
  with pd.ExcelWriter('结果.xlsx') as writer:
      df1.to_excel(writer, sheet_name='表1')
      df2.to_excel(writer, sheet_name='表2')
## Day 6（2026.03.10）
- **学习内容**：pandas 数据处理进阶（缺失值、重复值、异常值、数据类型转换、标准化、字符串操作）
- **练习文件**：[data_cleaning_advanced.py](./Day6/data_cleaning_advanced.py)
- **核心收获**：
  - 能够识别并处理财务数据中的缺失值、重复值、异常值
  - 掌握数据类型转换技巧，确保数据正确性
  - 学会数据标准化方法，为后续建模做准备
  - 能够清洗金蝶导出的科目余额表，为分析打基础
- **示例代码概览**：
  ```python
  # 处理缺失值
  df.fillna(df.mean(), inplace=True)

  # 检测异常值（IQR法）
  Q1 = df['金额'].quantile(0.25)
  Q3 = df['金额'].quantile(0.75)
  IQR = Q3 - Q1
  outliers = df[(df['金额'] < Q1-1.5*IQR) | (df['金额'] > Q3+1.5*IQR)]

  # 标准化
  df['金额_标准化'] = (df['金额'] - df['金额'].min()) / (df['金额'].max() - df['金额'].min())
