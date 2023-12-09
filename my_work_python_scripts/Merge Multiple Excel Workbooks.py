import os
import pandas as pd

"""
用来合并在一个目录中，多个格式和字段相同的Excel工作簿
"""

# 先设置好存放Excel文件的文件夹路径
folder_path = 'C:\\Users\\Administrator.DESKTOP-D6NNI6Q\\Desktop\\xunjian'
output_path = os.path.join(folder_path, '合并文件夹.xlsx')

# 初始化一个空的DateFrame 用于存储合并后的数据
merged_date = pd.DataFrame()

# 遍历文件夹中的每个Excel文件
for filename in os.listdir(folder_path):
    if filename.endswith(".xlsx"):  # 确保只读取后缀是“xlsx”的Excel文件
        file_path = os.path.join(folder_path, filename)

        # 读取Excel文件到DataFrame
        df = pd.read_excel(file_path)

        # 合并数据到主DataFrame
        merged_date = pd.concat([merged_date, df], ignore_index=True)

        # 将合并后的数据保存到新的Excel文件中
        merged_date.to_excel(output_path, index=False)
