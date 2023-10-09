import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog
from tkinter import ttk
from ttkthemes import ThemedStyle  # 导入ThemedStyle
from threading import Thread
import win32com.client as win32

# 创建主窗口
root = tk.Tk()
root.title("CSV to Excel Converter")

# 创建并应用主题样式
style = ThemedStyle(root)
style.set_theme("arc")  # 选择一个主题，例如"arc"


# 创建函数来转换CSV文件为Excel文件
def convert_csv_to_excel():
    input_folder = input_folder_var.get()
    output_folder = output_folder_var.get()
    vba_macro_file = vba_macro_entry.get()

    # 获取CSV文件列表
    csv_files = [f for f in os.listdir(input_folder) if f.endswith('.csv')]

    progress_bar['maximum'] = len(csv_files)

    for i, csv_file in enumerate(csv_files):
        csv_path = os.path.join(input_folder, csv_file)
        df = pd.read_csv(csv_path)
        excel_file = os.path.splitext(csv_file)[0] + '.xlsx'
        excel_path = os.path.join(output_folder, excel_file)
        df.to_excel(excel_path, index=False)
        progress_bar['value'] = i + 1
        root.update()
        # 在转换后的Excel文件中添加VBA宏并运行
        add_and_run_vba_macro(excel_path, vba_macro_file)

    progress_label.config(text="转换完成！")


def add_and_run_vba_macro(excel_file_path, vba_macro_file):
    # 创建一个 Excel 实例

    excel = win32.Dispatch("Excel.Application")

    excel.Visible = True  # 如果需要显示 Excel 界面，可以设置为 True

    # 打开 Excel 工作簿
    workbook = excel.Workbooks.Open(excel_file_path)  # 替换为您的工作簿路径

    # 导入 VBA 宏文件
    vba_module = workbook.VBProject.VBComponents.Import(vba_macro_file)  # 替换为您的宏文件路径

    # 运行 VBA 宏
    excel.Application.Run("CreateScatterChartWithLine")  # 替换为您的宏的名称

    # 保存工作簿
    workbook.Save()

    # 关闭工作簿和 Excel
    workbook.Close()
    excel.Quit()


def browse_input_folder():
    folder = filedialog.askdirectory()
    if folder:
        input_folder_var.set(folder)


def browse_output_folder():
    folder = filedialog.askdirectory()
    if folder:
        output_folder_var.set(folder)


def browse_vba_macro():
    file_path = filedialog.askopenfilename(filetypes=[("VBA Macro Files", "*.bas")])
    vba_macro_entry.delete(0, tk.END)
    vba_macro_entry.insert(0, file_path)


# 创建文件选择按钮和标签
input_folder_label = ttk.Label(root, text="选择CSV文件夹:")
input_folder_label.grid(row=0, column=0, padx=10, pady=10, sticky="w")
input_folder_var = tk.StringVar()
input_folder_entry = ttk.Entry(root, textvariable=input_folder_var, state="readonly")
input_folder_entry.grid(row=0, column=1, padx=10, pady=10, sticky="ew")
input_folder_button = ttk.Button(root, text="浏览", command=browse_input_folder)
input_folder_button.grid(row=0, column=2, padx=5, pady=10)

output_folder_label = ttk.Label(root, text="选择Excel输出文件夹:")
output_folder_label.grid(row=1, column=0, padx=10, pady=10, sticky="w")
output_folder_var = tk.StringVar()
output_folder_entry = ttk.Entry(root, textvariable=output_folder_var, state="readonly")
output_folder_entry.grid(row=1, column=1, padx=10, pady=10, sticky="ew")
output_folder_button = ttk.Button(root, text="浏览", command=browse_output_folder)
output_folder_button.grid(row=1, column=2, padx=5, pady=10)

vba_macro_label = ttk.Label(root, text="选择VBA宏文件：")
vba_macro_label.grid(row=2, column=0)
vba_macro_var = tk.StringVar()
vba_macro_entry = ttk.Entry(root, textvariable=vba_macro_var, width=40)
vba_macro_entry.grid(row=2, column=1, padx=10)
vba_macro_button = ttk.Button(root, text="浏览", command=browse_vba_macro)
vba_macro_button.grid(row=2, column=2)


# 创建转换按钮和进度条
convert_button = ttk.Button(root, text="开始转换", command=lambda: convert_csv_to_excel())
convert_button.grid(row=3, column=0, columnspan=3, pady=20)

progress_label = ttk.Label(root, text="")
progress_label.grid(row=4, column=0, columnspan=3)

progress_bar = ttk.Progressbar(root, orient="horizontal", length=300, mode="determinate")
progress_bar.grid(row=5, column=0, columnspan=3, pady=10)

root.mainloop()
