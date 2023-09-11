from tkinter import filedialog, Tk, Button, Label, Entry
from pandas import read_excel, isna, Series, ExcelWriter


def add_timecodes(timecode1, timecode2):
    def timecode_to_seconds(timecode):
        minutes, seconds, milliseconds = map(int, str(timecode).split('.'))
        total_seconds = minutes * 60 + seconds + milliseconds / 1000

        # 判断Total_Seconds的正负号
        if timecode.startswith('-'):
            total_seconds *= -1
        return total_seconds

    def seconds_to_timecode(seconds):
        minutes = int(seconds // 60)
        remaining_seconds = seconds % 60
        seconds = int(remaining_seconds)
        milliseconds = round((remaining_seconds - seconds) * 1000, 3)  # 修改这里，保留3位小数
        timecode = f"{minutes}.{seconds:02d}.{milliseconds:03.0f}"  # 修改这里，格式化为3位小数
        return timecode

    seconds1 = timecode_to_seconds(timecode1)
    seconds2 = timecode_to_seconds(timecode2)
    total_seconds = seconds1 + seconds2
    result_timecode = seconds_to_timecode(total_seconds)

    return result_timecode


def read_file(file_path, sheet_name):
    if file_path.endswith(('.xlsx', '.xls')):
        data = read_excel(file_path, sheet_name=sheet_name)
        print(file_path)
    else:
        raise ValueError("Unsupported file format.")

    col1_data = data['时间码'].tolist()
    col2_data = data.iloc[:, 5].tolist()
    length = len(col1_data)
    print("列表有：{}行".format(length))

    for i, value in enumerate(col2_data):
        # print(i,value)
        # 返回第二列的下标与数值
        if not isna(value):
            # print(value)
            for j in range(i, length):
                # print(j)
                print("时间码：{}".format(col1_data[j]))
                print("偏移时间：{}".format(str(value)))
                col1_data[j] = add_timecodes(col1_data[j], str(value))
    data["最终时间"] = Series(col1_data)
    data.iloc[:, 6] = Series(col1_data)

    with ExcelWriter(file_path) as writer:
        data.to_excel(writer, sheet_name=sheet_name, index=False)

    return col1_data, col2_data,


# D:\Python\Lib\site-packages
# E:\TimeOffset Calculator for Multilingual\TOCM.py
# pyinstaller -F  "E:\TimeOffset Calculator for Multilingual\TOCM.py" -p "D:\Python\Lib\site-packages"

def open_file():
    file_path = filedialog.askopenfilename(title="请选中带时间码的文件",
                                           filetypes=(("Excel Files", "*.xlsx;*.xls"),
                                                      ("All Files", "*.*")))
    print(file_path)
    col1_name = col1_name_entry.get()
    col2_name = col2_name_entry.get()
    if file_path:
        try:
            sheet_name = sheet_name_entry.get()
            print("SheetName:{}".format(sheet_name))
            col1_data, col2_data = read_file(file_path, sheet_name)
            print("Column 1 data:", col1_data)
            print("Column 2 data:", col2_data)
            from tkinter.messagebox import showinfo
            showinfo(title="Excel处理已完成", message="Excel文件已计算最终时间，请打开文件确认")
        except Exception as e:
            print(f"Error: {e}")


# 创建Tkinter界面
root = Tk()
root.title("Excel处理")

# 创建按钮，点击后执行open_file函数
open_button = Button(root, text="打开文件", command=open_file)
open_button.grid(row=0, column=0, padx=5, pady=5)

sheet_name_label = Label(root, text="Sheet名称")
sheet_name_label.grid(row=1, column=0, padx=5, pady=5)

# 创建输入框，用于输入Sheet名称
sheet_name_entry = Entry(root)
sheet_name_entry.grid(row=2, column=0, padx=5, pady=5)

# 创建标签，显示“第一列名称”
col1_name_label = Label(root, text="第一列名称")
col1_name_label.grid(row=3, column=0, padx=5, pady=5)

# 创建输入框，用于输入第一列名称
col1_name_entry = Entry(root)
col1_name_entry.grid(row=4, column=0, padx=5, pady=5)
col1_name_entry.insert(0, "时间码")  # 默认值

# 创建标签，显示“第二列名称”
col2_name_label = Label(root, text="第二列名称")
col2_name_label.grid(row=5, column=0, padx=5, pady=5)

# 创建输入框，用于输入第二列名称
col2_name_entry = Entry(root)
col2_name_entry.grid(row=6, column=0, padx=5, pady=5)
col2_name_entry.insert(0, "时间偏差")  # 默认值

# 运行Tkinter界面
root.mainloop()
