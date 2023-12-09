import pandas as pd

"""
用来把一个Excel工作簿，按照工作表的个数和名字，拆分成对应个数的工作簿，并命名
"""

# 读取原始工作簿
file_path = r'C:\Users\Administrator.DESKTOP-D6NNI6Q\Desktop\巡检工作\健康巡检异常项分析统计追踪表.xlsx'
original_workbook = pd.ExcelFile(file_path)

# 获取原始工作簿的所有工作表名
sheet_names = original_workbook.sheet_names

# 循环拆分并保存每个工作表为一个新工作簿
for sheet_name in sheet_names:
    # 读取原始工作表
    df = original_workbook.parse(sheet_name)

    # 获取当前日期
    current_date = pd.to_datetime('today').strftime('%Y%m%d')

    # 构建新工作簿的文件路径和名称
    new_workbook_name = f'C:\\Users\\Administrator.DESKTOP-D6NNI6Q\\Desktop\\巡检工作\\{current_date}_{sheet_name}.xlsx'

    # 将当前工作表保存为新工作簿
    df.to_excel(new_workbook_name, index=False)

print("拆分完成！")

