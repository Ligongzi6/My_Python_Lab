import re

import openpyxl
import pandas as pd

"""
这个python文件是为了生成当天巡检的工作量统计情况。
"""

# 设置要统计数量的日期，注意替换为实际的当前时间
current_time = input("请输入当前时间（格式为YYYY/MM/DD）：")

# Excel文件路径,这个路径最好就固定
file_path = r'C:\Users\Administrator.DESKTOP-D6NNI6Q\Desktop\巡检工作\健康巡检异常项分析统计追踪表.xlsx'

#  设置好要读入哪些工作表，所以总表中的工作表表名最好不要经常去改
sheets_to_read = [
    'Sheet1',
    'Sheet2',
    'Sheet3',
    'Sheet4',
    'Sheet5',
    'Sheet6',
    'Sheet7',
    'Sheet8',
    'Sheet9',
    'Sheet10'
    ]
# 开始读入上面对应列表中的工作表
all_sheets = pd.read_excel(file_path, sheet_name=sheets_to_read)


# # sheet_name=None，表示读取Excel文件中的所有工作表，
# all_sheets = pd.read_excel(file_path, sheet_name=None)

# 这里是事先定义好一个拿来存储结果的列表
result_list = []

# 遍历每个工作表
for sheet_name, sheet_data in all_sheets.items():
    # 统计分析时间列的值为当前时间的行数
    analysis_time_count = sheet_data[sheet_data['分析时间'] == current_time].shape[0]

    # 统计提单时间列的值为当前时间的行数
    submit_time_count = sheet_data[sheet_data['提单时间'] == current_time].shape[0]

    # 统计归档时间列的值为当前时间的行数
    archive_time_count = sheet_data[sheet_data['归档时间'] == current_time].shape[0]

    # 统计关闭提单列的值为当前时间的行数
    close_submit_count = sheet_data[sheet_data['关闭提单'] == current_time].shape[0]

    # 输出统计结果，这段删掉也没关系
    print(f"应用：{sheet_name}")
    print(f"分析异常项个数：{analysis_time_count}")
    print(f"提交缺陷单个数：{submit_time_count}")
    print(f"归档异常单个数：{archive_time_count}")
    print(f"关闭缺陷单个数：{close_submit_count}")
    print("\n")

    # 将结果添加到列表中
    result_list.append({
        '应用': sheet_name,
        '分析异常项个数': analysis_time_count,
        '提交缺陷单个数': submit_time_count,
        '归档异常单个数': archive_time_count,
        '关闭缺陷单个数': close_submit_count
    })

# 将结果列表转换为DataFrame
result_df = pd.DataFrame(result_list)

# 在DataFrame中添加一行进行汇总
result_df.loc['汇总'] = result_df.iloc[:, 1:].sum()

# 处理NaN和INF值，将它们替换为空字符串
result_df = result_df.replace({pd.NA: '', pd.NaT: '', float('inf'): '', float('-inf'): ''})

# 手动在“汇总”行的第一个单元格内填入“汇总”字样
result_df.iloc[-1, 0] = '汇总'

# 将DataFrame保存到新的Excel文件
# 至此统计完成
output_file_path = r'C:\Users\Administrator.DESKTOP-D6NNI6Q\Desktop\巡检工作量统计.xlsx'
result_df.to_excel(output_file_path, index=False)

# 从这里往下是调整上面统计完成而生成的表格的格式的
# 打开工作簿并选择工作表
workbook = openpyxl.load_workbook(output_file_path)
# 获取当前活动的工作表,本来就一个表，这么写没关系
worksheet = workbook.active

# 遍历每个列并设置最合适的宽度
for col in worksheet.columns:
    max_length = max([
        len(str(cell.value)) + 0.7 * len(re.findall(r'([\u4e00-\u9fa5])', str(cell.value)))
        for cell in col
    ])
    worksheet.column_dimensions[col[0].column_letter].width = (max_length + 2) * 1.2

# 保存工作簿
workbook.save(output_file_path)

print(f"结果已保存到：{output_file_path}")
