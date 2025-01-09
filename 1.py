import pandas as pd
import os
import re
import tkinter as tk
from tkinter import filedialog, messagebox
from openpyxl.styles import Alignment
from openpyxl.worksheet.datavalidation import DataValidation  # 导入DataValidation类

# 创建Tkinter根窗口并隐藏
root = tk.Tk()
root.withdraw()

# 选择输入文件夹路径
input_folder_path = filedialog.askdirectory(title="选择输入文件夹")
if not input_folder_path:
    print("未选择输入文件夹，程序结束。")
    exit()

# 设置输出文件路径为桌面
output_file_path = os.path.join(os.path.expanduser("~"), "Desktop", "export.xlsx")

# 创建输出文件夹（如果不存在）
output_folder = os.path.dirname(output_file_path)
os.makedirs(output_folder, exist_ok=True)

# 定义要删除的文件名
files_to_delete = ['index.xls', 'index.xlsx']
for file_name in files_to_delete:
    file_path = os.path.join(input_folder_path, file_name)
    if os.path.isfile(file_path):
        confirmation = messagebox.askyesno("确认删除", f"确认要删除文件 {file_path} 吗？")
        if confirmation:
            os.remove(file_path)
            print(f"已删除: {file_path}")
        else:
            print(f"未删除: {file_path}")
    else:
        print(f"文件未找到: {file_path}")

# 定义要检查的高危端口
high_ports = [135, 136, 137, 139, 445]

# 使用with语句来确保ExcelWriter自动关闭
with pd.ExcelWriter(output_file_path, engine='openpyxl') as writer:
    # 获取所有Excel文件
    excel_files = [f for f in os.listdir(input_folder_path) if f.endswith(('.xls', '.xlsx'))]
    total_files = len(excel_files)

    for idx, filename in enumerate(excel_files, start=1):
        print(f"正在处理文件 {idx}/{total_files}: {filename}")

        try:
            # 读取“远程漏洞”和“其它信息”工作表
            remote_vuln_df = pd.read_excel(os.path.join(input_folder_path, filename), sheet_name='远程漏洞')
            other_info_df = pd.read_excel(os.path.join(input_folder_path, filename), sheet_name='其它信息')

            # 检查是否有足够的列
            if remote_vuln_df.shape[1] < 6:
                print(f"警告: {filename} 列数不足，跳过此文件。")
                continue

            # 提取D列和F列的所有值，跳过标题行
            d_column = remote_vuln_df.iloc[1:, 3]  # D列（第四列）
            f_column = remote_vuln_df.iloc[1:, 5]  # F列（第六列）

            # 获取不带扩展名的文件名
            base_filename = os.path.splitext(filename)[0]

            # 创建一个新的DataFrame来保存提取的结果
            result_df = pd.DataFrame({
                'IP': [base_filename] * len(d_column),
                '漏洞名称': d_column,
                '风险等级': f_column
            })

            # 在“风险等级”后插入三列
            result_df.insert(result_df.columns.get_loc('风险等级') + 1, '是否整改', '')
            result_df.insert(result_df.columns.get_loc('风险等级') + 2, '整改措施', '')
            result_df.insert(result_df.columns.get_loc('风险等级') + 3, '高危端口', '')
            result_df.insert(result_df.columns.get_loc('风险等级') + 4, '备注', '')

            # 删除所有列中的方括号
            result_df = result_df.replace(r'\[|\]', '', regex=True)

            # 确保 B 列存在
            if other_info_df.shape[1] < 2:
                print("警告: '其它信息'工作表中没有足够的列。")
                continue

            other_info_column = other_info_df.iloc[:, 1]  # 使用 iloc 获取第二列（B 列）
            found_ports = set()  # 使用集合来存储找到的端口，自动去重

            # 提取所有可能的端口
            for info in other_info_column.dropna():
                extracted_ports = re.findall(r'\b\d+\b', str(info))
                for port in extracted_ports:
                    if int(port) in high_ports:
                        found_ports.add(int(port))

            # 筛选出风险等级中不包含"[低]"的项
            high_mid_df = result_df[~result_df['风险等级'].str.contains(r'\[低\]|低', na=False)]

            # 将找到的高危端口写入到结果DataFrame中，并去重
            if found_ports:
                unique_ports = sorted(set(found_ports))  # 去重并排序
                high_mid_df.loc[:, '高危端口'] = ', '.join(map(str, unique_ports))
            else:
                high_mid_df.loc[:, '高危端口'] = ''  # 确保列名一致

            # 检查除了A1以外的行是否为空
            if high_mid_df.iloc[:, 1:].dropna(how='all').empty:
                print(f"工作表 '{base_filename}' 没有有效数据，跳过写入。")
                continue  # 跳过写入该工作表

            # 更新结果DataFrame
            result_df = high_mid_df
            result_df.insert(0, '序号', range(1, len(result_df) + 1))  # 重新生成“序号”列

            # 将结果写入到Excel文件的一个工作表中，以文件名作为工作表名
            sheet_name = base_filename[:31]  # Excel工作表名称最大长度为31
            result_df.to_excel(writer, sheet_name=sheet_name, index=False)

            # 获取工作表对象
            worksheet = writer.sheets[sheet_name]

            # 自动调整列宽
            for i, col in enumerate(result_df.columns):
                max_len = max(result_df[col].astype(str).map(len).max(), len(col))
                worksheet.column_dimensions[worksheet.cell(row=1, column=i + 1).column_letter].width = max_len + 5

            # 设置列A、B、D、E和G的居中对齐
            columns_to_center = ['A', 'B', 'D', 'E', 'G']
            for col in columns_to_center:
                for row in range(1, len(result_df) + 2):  # 包括标题行
                    worksheet[f'{col}{row}'].alignment = Alignment(horizontal='center')

            # # 为E列添加下拉列表数据验证，并设置输入提示
            # dv = DataValidation(type="list", formula1='"是,否"', showDropDown=True)

            # # 设置输入提示信息
            # dv.showInputMessage = True
            # dv.promptTitle = "整改提示"
            # dv.prompt = "请选择'是'后，请输入整改过程或截图。"

            worksheet.add_data_validation(dv)

            # 应用数据验证到E列
            for row in range(2, len(result_df) + 2):  # 从第二行开始（跳过标题行）
                dv.add(worksheet[f'E{row}'])


            print(f"已处理 {filename}: 数据已保存到工作表 '{sheet_name}'")
        except Exception as e:
            print(f"处理 {filename} 时出错: {e}")

# 输出文件路径提示
print(f"所有文件处理完成，输出文件路径: {output_file_path}")
