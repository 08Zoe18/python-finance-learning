import pandas as pd
import os

# ================== 配置区域 ==================
# 请修改为你的文件实际路径（从金蝶导出的明细账Excel）
file_path = r"F:\Desktop\Excel练习\江右-3月财务分析.xlsx"
# ==============================================

# 检查文件是否存在（如果不存在会提示）
if not os.path.exists(file_path):
    print(f"错误：文件不存在，请检查路径：{file_path}")
    exit()

# 1. 读取Excel文件
df = pd.read_excel(file_path)

# 2. 查看列名，确认你的表格里实际有哪些列
print("列名如下：")
print(df.columns.tolist())

# 3. 根据你的实际列名设置（从截图看，列名是“科目名称”、“借方”、“贷方”）
科目列 = '科目名称'
借方列 = '借方'
贷方列 = '贷方'

# 4. 打印所有科目名称，以便确认产品名称的关键词（方便调试）
print("\n所有科目名称（唯一值）：")
print(df[科目列].unique())

# 5. 筛选各产品收入（贷方发生额）
# 注意：科目名称中可能包含“A产品”、“B产品”等关键词，我们用 `.str.contains()` 匹配
income_a = df[df[科目列].str.contains('A产品', na=False)][贷方列].sum()
income_b = df[df[科目列].str.contains('B产品', na=False)][贷方列].sum()
income_c = df[df[科目列].str.contains('C产品', na=False)][贷方列].sum()

# 6. 筛选各产品成本（借方发生额）
cost_a = df[df[科目列].str.contains('A产品', na=False)][借方列].sum()
cost_b = df[df[科目列].str.contains('B产品', na=False)][借方列].sum()
cost_c = df[df[科目列].str.contains('C产品', na=False)][借方列].sum()

# 7. 计算毛利和毛利率
profit_a = income_a - cost_a
profit_b = income_b - cost_b
profit_c = income_c - cost_c

gross_margin_a = profit_a / income_a if income_a != 0 else 0
gross_margin_b = profit_b / income_b if income_b != 0 else 0
gross_margin_c = profit_c / income_c if income_c != 0 else 0

# 8. 打印结果
print("\n=== 3月各产品分析结果 ===")
print(f"A产品：收入 {income_a:,.2f}，成本 {cost_a:,.2f}，毛利 {profit_a:,.2f}，毛利率 {gross_margin_a:.2%}")
print(f"B产品：收入 {income_b:,.2f}，成本 {cost_b:,.2f}，毛利 {profit_b:,.2f}，毛利率 {gross_margin_b:.2%}")
print(f"C产品：收入 {income_c:,.2f}，成本 {cost_c:,.2f}，毛利 {profit_c:,.2f}，毛利率 {gross_margin_c:.2%}")

# 9. 保存结果到新Excel文件
summary = pd.DataFrame({
    '产品': ['A', 'B', 'C'],
    '收入': [income_a, income_b, income_c],
    '成本': [cost_a, cost_b, cost_c],
    '毛利': [profit_a, profit_b, profit_c],
    '毛利率': [gross_margin_a, gross_margin_b, gross_margin_c]
})

output_file = '3月产品分析_Python结果.xlsx'
summary.to_excel(output_file, index=False)
print(f"\n结果已保存到文件：{output_file}")