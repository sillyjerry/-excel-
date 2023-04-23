# import os
# import pandas as pd
#
# # 定义文件夹路径
# folder_path = r'C:\Users\86159\Desktop\算账\2204-2301算账表\23年一季度算账表\2023年第一季度脱贫户收入算账表'  # 请将文件夹路径替换成您的实际文件夹路径
#
# # 获取文件夹内所有文件的文件名
# files = os.listdir(folder_path)
#
# # 按照文件名排序
# files.sort()
#
# # 创建一个空的 DataFrame 用于存储合并后的数据
# merged_data = pd.DataFrame()
#
# # 循环读取并合并所有 Excel 文件
# for file in files:
#     if file.endswith('.xlsx') or file.endswith('.xls'):  # 确保文件是 Excel 文件
#         file_path = os.path.join(folder_path, file)
#         data = pd.read_excel(file_path)  # 使用 pandas 读取 Excel 文件
#
#         # 提取固定位置的数据并添加到合并后的数据中
#         extracted_data = data.iloc[2, 2]  # 请将固定行号和固定列号替换成实际的行号和列号
#         extracted_data_series = pd.Series([extracted_data], name=file)  # 将单值转换为 Series，并设置 Series 名称为文件名
#         merged_data = pd.concat([merged_data, extracted_data_series], axis=1, ignore_index=True)  # 添加提取的数据到 DataFrame
#
# # 输出合并后的数据
# print(merged_data)
# merged_data.to_excel('output.xlsx', index=False)
#
##完美版本
# import os
# import pandas as pd
#
# # 定义文件夹路径
# folder_path = r'C:\Users\86159\Desktop\算账\2204-2301算账表\23年一季度算账表\2023年第一季度脱贫户收入算账表'  # 请将文件夹路径替换成您的实际文件夹路径
#
# # 获取文件夹内所有文件的文件名
# files = os.listdir(folder_path)
#
# # 按照文件名排序
# files.sort()
#
# # 创建一个空的 DataFrame 用于存储合并后的数据
# merged_data = pd.DataFrame()
#
# # 定义需要提取的单元格位置
# cell_locations = {'C4': (2, 2), 'C5': (3, 2), 'C9': (7, 2), 'C16': (14, 2), 'F8': (6, 5)}
#
# # 循环读取并合并所有 Excel 文件
# for file in files:
#     if file.endswith('.xlsx') or file.endswith('.xls'):  # 确保文件是 Excel 文件
#         file_path = os.path.join(folder_path, file)
#         data = pd.read_excel(file_path)  # 使用 pandas 读取 Excel 文件
#
#         # 提取固定位置的数据并添加到合并后的数据中
#         extracted_data = []
#         for location in cell_locations.values():
#             extracted_data.append(data.iloc[location[0], location[1]])
#         extracted_data_series = pd.Series(extracted_data, name=file)  # 将提取的数据转换为 Series，并设置 Series 名称为文件名
#         merged_data = pd.concat([merged_data, extracted_data_series], axis=1)  # 添加提取的数据到 DataFrame
#
# # 使用文件名作为索引列
# merged_data.index = cell_locations.keys()
#
# # 输出合并后的数据
# print(merged_data)
# merged_data.to_excel('output.xlsx', index_label='File Name')


##排序
# import os
# import pandas as pd
#
# # 定义文件夹路径
# folder_path = r'C:\Users\86159\Desktop\算账\2204-2301算账表\23年一季度算账表\2023年第一季度脱贫户收入算账表'  # 请将文件夹路径替换成您的实际文件夹路径
#
# # 获取文件夹内所有文件的文件名
# files = os.listdir(folder_path)
#
# # 按照文件名排序
# files.sort()
#
# # 创建一个空的 DataFrame 用于存储合并后的数据
# merged_data = pd.DataFrame()
#
# # 定义需要提取的单元格位置和对应的行索引名称
# cell_locations = {'C4': (2, 2, '务工收入'), 'C5': (3, 2, '生产经营性收入'), 'C9': (7, 2, '财产性收入'),
#                   'C16': (14, 2, '各项补贴'), 'F8': (6, 5, '生产经营性支出')}
#
# # 循环读取并合并所有 Excel 文件
# for file in files:
#     if file.endswith('.xlsx') or file.endswith('.xls'):  # 确保文件是 Excel 文件
#         file_path = os.path.join(folder_path, file)
#         data = pd.read_excel(file_path)  # 使用 pandas 读取 Excel 文件
#
#         # 提取固定位置的数据并添加到合并后的数据中
#         extracted_data = []
#         for location in cell_locations.values():
#             extracted_data.append(data.iloc[location[0], location[1]])
#         extracted_data_series = pd.Series(extracted_data, name=file)  # 使用文件名作为列索引
#         merged_data = pd.concat([merged_data, extracted_data_series], axis=1)  # 添加提取的数据到 DataFrame
#
# # 将行索引设置为指定的名称
# merged_data.index = [cell_locations[key][2] for key in cell_locations]
#
# # 输出合并后的数据
# print(merged_data)
# merged_data.to_excel('output.xlsx', index_label='行索引名称')


##不排序
import os
import pandas as pd

# 定义文件夹路径
folder_path = r'C:\Users\86159\Desktop\算账\2204-2301算账表\23年一季度算账表\2023年第一季度脱贫户收入算账表'  # 请将文件夹路径替换成您的实际文件夹路径

# 获取文件夹内所有文件的文件名
files = os.listdir(folder_path)

# 按照文件名排序
##files.sort()

# 创建一个空的 DataFrame 用于存储合并后的数据
merged_data = pd.DataFrame()

# 定义需要提取的单元格位置和对应的行索引名称
cell_locations = {'C4': (2, 2, '务工收入'), 'C5': (3, 2, '生产经营性收入'), 'C9': (7, 2, '财产性收入'),
                  'C16': (14, 2, '各项补贴'), 'F8': (6, 5, '生产经营性支出')}

# 循环读取并合并所有 Excel 文件
for file in files:
    if file.endswith('.xlsx') or file.endswith('.xls'):  # 确保文件是 Excel 文件
        file_path = os.path.join(folder_path, file)
        data = pd.read_excel(file_path)  # 使用 pandas 读取 Excel 文件

        # 提取固定位置的数据并添加到合并后的数据中
        extracted_data = []
        for location in cell_locations.values():
            extracted_data.append(data.iloc[location[0], location[1]])
        extracted_data_series = pd.Series(extracted_data, name=file)  # 使用文件名作为列索引
        merged_data = pd.concat([merged_data, extracted_data_series], axis=1)  # 添加提取的数据到 DataFrame

# 将行索引设置为指定的名称
merged_data.index = [cell_locations[key][2] for key in cell_locations]

# 输出合并后的数据
print(merged_data)
merged_data.to_excel('output.xlsx', index_label='行索引名称')
