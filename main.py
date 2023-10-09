import win32com.client as win32

# 创建一个 Excel 实例
excel = win32.Dispatch("Excel.Application")
excel.Visible = True  # 如果需要显示 Excel 界面，可以设置为 True

# 打开 Excel 工作簿
workbook = excel.Workbooks.Open(r"F:\OperateExcel\2023-09-07 像素#7.xlsx")  # 替换为您的工作簿路径

# 导入 VBA 宏文件
vba_module = workbook.VBProject.VBComponents.Import(r"F:\OperateExcel\CreateChart.bas")  # 替换为您的宏文件路径

# 运行 VBA 宏
excel.Application.Run("CreateScatterPlotWithLine")  # 替换为您的宏的名称

# 保存工作簿
workbook.Save()

# 关闭工作簿和 Excel
workbook.Close()
excel.Quit()
