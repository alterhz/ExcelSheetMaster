import os
import win32com.client
import logging


def get_filename_from_path(path):
    return os.path.basename(path)


def open_excel_sheet(file_path, sheet_name):
    # 先判断文件是否存在
    if not os.path.exists(file_path):
        print(f"文件 {file_path} 不存在。")
        return None
    excel = win32com.client.Dispatch("Excel.Application")
    # 判断文件是否已经打开
    opened_workbooks = excel.Workbooks
    workbook_opened = False
    for workbook in opened_workbooks:
        if get_filename_from_path(workbook.FullName) == get_filename_from_path(file_path):
            workbook_opened = True
            break
    if not workbook_opened:
        workbook = excel.Workbooks.Open(file_path)
        logging.debug("打开文件:{file_path}")
    worksheet = None
    for sheet in workbook.Worksheets:
        if sheet.Name == sheet_name:
            worksheet = sheet
            break
    if worksheet is None:
        # 如果指定的工作表不存在，只打开文件
        print(f"工作表 {sheet_name} 不存在，仅打开文件。")
    else:
        # 如果工作表存在，激活该工作表
        excel.Visible = True
        worksheet.Activate()
        logging.debug("激活工作表:{sheet_name}")
        # 设置第一行第二列的值为“name”
        worksheet.Cells(1, 2).Value = "name"
    return worksheet


def test_open_excel_sheet():
    current_directory = os.path.dirname(os.path.abspath(__file__))
    file_path = os.path.join(current_directory, 'cache.xlsx')
    sheet_name = "config"
    worksheet = open_excel_sheet(file_path, sheet_name)
    if worksheet is None:
        print(f"无法打开文件 {file_path} 中的工作表 {sheet_name}。")
    else:
        print(f"成功打开文件 {file_path} 中的工作表 {sheet_name}。")


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO)
    test_open_excel_sheet()
