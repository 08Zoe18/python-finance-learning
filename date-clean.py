# =====================================================
# 销售数据清洗脚本
# 功能：删除重复行 + 按产品分类汇总金额
# 依赖库：pandas, openpyxl
# 使用方法：修改 file_path 为你的文件路径
# =====================================================

import pandas as pd

# ---------- 配置区域 ----------
file_path = "销售数据.xlsx"          # 你的 Excel 文件（放在同一文件夹）
sheet_name = "Sheet1"                # 工作表名称
product_col = "产品名称"              # 产品列名（根据你的文件修改）
amount_col = "金额"                   # 金额列名（根据你的文件修改）
# --------------------------------

# 1. 读取数据
df = pd.read_excel(file_path, sheet_name=sheet_name)

print("原始数据行数：", len(df))

# 2. 删除完全重复的行
df_clean = df.drop_duplicates()
print("删除重复后行数：", len(df_clean))

# 3. 按产品分类汇总金额
summary = df_clean.groupby(product_col)[amount_col].sum().reset_index()
print("\n各产品销售额汇总：")
print(summary)

# 4. 保存清洗后的数据
summary.to_excel("销售汇总_清洗后.xlsx", index=False)
print("\n汇总结果已保存到：销售汇总_清洗后.xlsx")