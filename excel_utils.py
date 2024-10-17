import logging
import os

import openpyxl
import pandas as pd
import win32com.client

from os_utils import get_filename_from_path


def get_sheet_names_fast(file_path):
    try:
        xls = pd.ExcelFile(file_path)
        sheet_names = xls.sheet_names
        xls.close()
        return sheet_names
    except Exception as e:
        return f"发生错误：{e}"


def get_excel_sheet_names(file_name):
    try:
        workbook = openpyxl.load_workbook(file_name)
        sheet_names = workbook.sheetnames
        workbook.close()
        return sheet_names
    except Exception as e:
        return f"发生错误：{e}"


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
    return worksheet


if __name__ == '__main__':
    # 获取当前脚本文件所在的目录
    current_directory = os.path.dirname(os.path.abspath(__file__))
    file_path = os.path.join(current_directory, 'cache.xlsx')
    sheet_name = '缓存文件列表'
    open_excel_sheet(file_path, sheet_name)
