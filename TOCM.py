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


def open_file():
    file_path = filedialog.askopenfilename(title="请选中带时间码的文件",
                                           filetypes=(("Excel Files", "*.xlsx;*.xls"),
                                                      ("All Files", "*.*")))
    if file_path:
        try:
            sheet_name = sheet_name_entry.get()
            col1_data, col2_data = read_file(file_path, sheet_name)
            print("Column 1 data:", col1_data)
            print("Column 2 data:", col2_data)
        except Exception as e:
            print(f"Error: {e}")


root = tk.Tk()
root.title("文件阅读器")

open_button = tk.Button(root, text="打开文件", command=open_file)
open_button.pack()

tk.Label(root, text="Sheet名称：").pack()
sheet_name_entry = tk.Entry(root)
sheet_name_entry.pack()

root.mainloop()

# Column 1 data: ['0.12.200', '0.22.766', '0.26.466', '0.34.610', '0.37.910', '0.45.010', '0.48.376', '0.55.210', '1.08.444']
# Column 1 data: ['0.12.200', '0.22.766', '0.26.466', '0.34.610', '0.37.910', '0.45.010', '0.48.376', '0.56.210', '1.09.444']