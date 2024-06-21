import os

import pandas as pd
import numpy as np

from openpyxl import load_workbook

# # 定义除法运算函数
# def safe_divide(numerator, denominator):
#     return np.where(denominator == 0, '#DIV/0!', numerator / denominator)


# 读取Excel文件中的数据
file1 = 'data_volume_check/File1.xlsx'
file2 = 'data_volume_check/File2.xlsx'

# 假设Sheet1是工作表的名称
df1 = pd.read_excel(file1, sheet_name='C360-2.0')
df2 = pd.read_excel(file2, sheet_name='c360')

# 创建辅助列（拼接A列和B列）  pandas 默认读取第一行作为列名
df1['Key'] = df1['data_source_name'].astype(str) + df1['table_name'].astype(str)
df2['Key'] = df2['data_source_name'].astype(str) + df2['table_name'].astype(str)

# 选择需要合并的列，包括辅助列
df2_subset = df2[['Key', 'source_data_volumn', 'target_data_volumn']]

# 合并数据
merged_df = pd.merge(df1, df2_subset, on='Key', how='left')

# 删除辅助列
merged_df.drop(columns=['Key'], inplace=True)

# 将NaN替换为#N/A
merged_df = merged_df.fillna('#N/A')


# 定义一个函数来处理指定的列
# def process_specified_columns(df):
#     # 添加新列E和F，并初始化
#     df['Compare'] = 0
#     df['Increase%'] = 0
#
#     # 处理Compare
#     def process_Compare(row):
#         cur_col = -2  # 当前列
#         # if(safe_divide(row[cur_col-1]-row[cur_col-2]  , row[cur_col-1]-row[cur_col-2]) == 0):
#
#         return 1 if safe_divide(row[cur_col - 1] - row[cur_col - 3], row[cur_col - 3] - row[cur_col - 5]) > 1.5 else 0
#
#     # 处理Increase列
#     def process_Increase(row):
#         cur_col = -1  # 当前列
#         return (row[cur_col-2] - row[cur_col-4]) / row[cur_col-2]
#
#     # 应用函数到E列
#     df['Compare'] = df.apply(lambda row: process_Compare(row), axis=1)
#
#     # 应用函数到F列
#     df['Increase%'] = df.apply(lambda row: process_Increase(row), axis=1)
#
#     return df


# # 应用处理函数
# merged_df = process_specified_columns(merged_df)

# s输出前先对 列名更改一下

output_file = 'data_volume_check/Merged_File.xlsx'

# 如果文件存在，则删除它
if os.path.exists(output_file):
    os.remove(output_file)

# 保存结果到新的Excel文件
merged_df.to_excel(output_file, index=False)

# wb = load_workbook(output_file)
# ws = wb.active # 获取当前活动的工作表
#
# # 获取列数  也就是当前
# total_columns = ws.max_column