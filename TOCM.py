import tkinter as tk
from tkinter import filedialog
import pandas as pd
import numpy as np

def add_timecodes(timecode1, timecode2):
    def timecode_to_seconds(timecode):
        minutes, seconds, milliseconds = map(int, timecode.split('.'))
        total_seconds = minutes * 60 + seconds + milliseconds / 1000
        return total_seconds

    def seconds_to_timecode(seconds):
        minutes = int(seconds // 60)
        remaining_seconds = seconds % 60
        seconds = int(remaining_seconds)
        milliseconds = int((remaining_seconds - seconds) * 1000)
        timecode = f"{minutes}.{seconds:02d}.{milliseconds:03d}"
        return timecode

    seconds1 = timecode_to_seconds(timecode1)
    seconds2 = timecode_to_seconds(timecode2)
    total_seconds = seconds1 + seconds2
    result_timecode = seconds_to_timecode(total_seconds)

    return result_timecode


def read_file(file_path, sheet_name):
    if file_path.endswith(('.xlsx', '.xls')):
        data = pd.read_excel(file_path, sheet_name=sheet_name)
    else:
        raise ValueError("Unsupported file format.")

    col1_data = data['时间码'].tolist()
    col2_data = data.iloc[:, 5].tolist()
    length = len(col1_data)

    for i, value in enumerate(col2_data):
        # print(i,value)
        # 返回第二列的下标与数值
        if not pd.isna(value):
            # print(value)
            for j in range(i, length):
                # print(j)
                print("时间码：{}".format(col1_data[j]))
                print("偏移时间：{}".format(str(value)))
                col1_data[j] = add_timecodes(col1_data[j], str(value))
    data["最终时间"] = pd.Series(col1_data)
    data.iloc[:, 6] = pd.Series(col1_data)

    with pd.ExcelWriter(file_path) as writer:
        data.to_excel(writer, sheet_name=sheet_name, index=False)

    return col1_data, col2_data,

# D:\Python\Lib\site-packages
# E:\TimeOffset Calculator for Multilingual\TOCM.py
def open_file():
    file_path = filedialog.askopenfilename(title="请选中带时间码的文件",
                                           filetypes=(("Excel Files", "*.xlsx;*.xls"),
                                                      ("All Files", "*.*")))
    col1_name = col1_name_entry.get()
    col2_name = col2_name_entry.get()
    if file_path:
        try:
            sheet_name = sheet_name_entry.get()
            col1_data, col2_data = read_file(file_path, sheet_name)
            print("Column 1 data:", col1_data)
            print("Column 2 data:", col2_data)
        except Exception as e:
            print(f"Error: {e}")


# 创建Tkinter界面
root = tk.Tk()
root.title("Excel处理")

# 创建按钮，点击后执行open_file函数
open_button = tk.Button(root, text="打开文件", command=open_file)
open_button.grid(row=0, column=0, padx=5, pady=5)

sheet_name_label = tk.Label(root, text="Sheet名称")
sheet_name_label.grid(row=1, column=0, padx=5, pady=5)

# 创建输入框，用于输入Sheet名称
sheet_name_entry = tk.Entry(root)
sheet_name_entry.grid(row=2, column=0, padx=5, pady=5)


# 创建标签，显示“第一列名称”
col1_name_label = tk.Label(root, text="第一列名称")
col1_name_label.grid(row=3, column=0, padx=5, pady=5)

# 创建输入框，用于输入第一列名称
col1_name_entry = tk.Entry(root)
col1_name_entry.grid(row=4, column=0, padx=5, pady=5)
col1_name_entry.insert(0, "时间码")  # 默认值

# 创建标签，显示“第二列名称”
col2_name_label = tk.Label(root, text="第二列名称")
col2_name_label.grid(row=5, column=0, padx=5, pady=5)

# 创建输入框，用于输入第二列名称
col2_name_entry = tk.Entry(root)
col2_name_entry.grid(row=6, column=0, padx=5, pady=5)
col2_name_entry.insert(0, "时间偏差")  # 默认值

# 运行Tkinter界面
root.mainloop()