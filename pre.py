import tkinter as tk
from tkinter import messagebox

# 初始化字典，用于存储input.txt文件的第二列内容

# 读取input.txt文件

def conv_dict():
    input_dict = {}
    data_dict = {}

    with open('input.txt', 'r', encoding='utf-8') as file_input:
        for line_number, line in enumerate(file_input, start=1):  # 假设第一行是标题，从第二行开始读取
            parts = line.strip().split()
            if len(parts) >= 2:  # 确保行中有足够的数据
                input_dict[line_number] = parts[1]  # 假设第二列的数据是我们想要的

    # 读取data.txt文件并替换第二列
    with open('data.txt', 'r', encoding='utf-8') as file_data, \
        open('data_replaced.txt', 'w', encoding='utf-8') as file_output:  # 将替换结果写入新文件
        for line in file_data:
            parts = line.strip().split(';')
            if len(parts) == 2:
                col_identifier, num_str = parts
                try:
                    # 尝试从input_dict中获取对应的值
                    replacement = input_dict[int(num_str)]
                    # 写入新的data_dict字典
                    data_dict[col_identifier] = replacement
                    # 写入替换后的行到新文件
                    file_output.write(f"{col_identifier};{replacement}\n")
                except KeyError:
                    # 如果索引超出范围，则弹出消息框提示
                    tk.Tk().withdraw()  # 不显示主窗口
                    messagebox.showerror("索引超出范围", "data.txt中的某个索引在input.txt中不存在。")
                    break
                except ValueError:
                    # 如果转换为整数失败，则弹出消息框提示
                    tk.Tk().withdraw()  # 不显示主窗口
                    messagebox.showerror("值错误", "data.txt中的数字格式不正确。")
                    break
            else:
                # 如果data.txt文件的行格式不正确，则写入原始行到新文件
                file_output.write(line)
    return data_dict
# 打印data_dict字典以验证