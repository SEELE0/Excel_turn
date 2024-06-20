import pandas as pd

# 读取Excel文件
file_path = 'your_excel_file.xlsx'  # 替换为你的文件路径
df = pd.read_excel(file_path)

# 删除空白项
df_cleaned = df.dropna()

# 保存结果到新的Excel文件
output_file_path = 'cleaned_excel_file.xlsx'  # 替换为输出文件路径
df_cleaned.to_excel(output_file_path, index=False)

print(f"清理后的文件已保存到: {output_file_path}")
