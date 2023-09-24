import win32com.client as win32
# 从Excel excel_name 的sheet_name 中导出图片保存至picture_name 中
def export_picture(excel_name, picture_name, sheet_name):
    # 获取Excel api
    excel = win32.gencache.EnsureDispatch("Excel.Application")
    
    # 打开Excel 文档, wb 为文件句柄
    wb = excel.Workbooks.Open("excel_name.xlsx")
    
    # 导出图片
    wb.Sheets("sheet_name").Export("picture_name.jpg")
    
    wb.Close()
