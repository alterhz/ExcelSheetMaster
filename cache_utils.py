import logging
import os
import time

import openpyxl

from excel_sheet_handler import ExcelSheetHandler
from excel_utils import get_excel_sheet_names
from os_utils import get_current_file_names, get_filename_from_path

CACHE_EXCEL_NAME = "cache.xlsx"
CONFIG_SHEET_NAME = "config"
ALL_PATH_SHEET_NAME = "allPath"


def create_excel_sheet(excel_name, sheet_name=None):
    """
    检查缓存文件是否存在，如果不存在则创建，并可根据需求创建指定名称的工作表。

    参数：
    excel_name：Excel 文件名称。
    sheet_name（可选）：要创建的工作表名称，如果不提供则不进行创建工作表的操作。

    返回：
    bool：如果创建了文件或者页签返回 True，否则返回 False。
    """
    file_created = False
    sheet_created = False
    if not os.path.exists(excel_name):
        wb = openpyxl.Workbook()
        wb.save(excel_name)
        file_created = True
    workbook = openpyxl.load_workbook(excel_name)
    if sheet_name and sheet_name not in workbook.sheetnames:
        workbook.create_sheet(sheet_name)
        sheet_created = True
    workbook.save(excel_name)
    return file_created or sheet_created


def create_config_sheet_header():
    if create_excel_sheet(CACHE_EXCEL_NAME, CONFIG_SHEET_NAME):
        with ExcelSheetHandler(CACHE_EXCEL_NAME, CONFIG_SHEET_NAME) as handler_config:
            handler_config.insert_column_header(1, "cs", "String", "key", "键", 20)
            handler_config.insert_column_header(2, "cs", "String", "value", "值", 30)
            handler_config.save_workbook()


def set_config_value(key, value):
    with ExcelSheetHandler(CACHE_EXCEL_NAME, CONFIG_SHEET_NAME) as handler_config:
        data = handler_config.get_all_data()
        for row in data:
            if row["key"] == key:
                row["value"] = value
                handler_config.write_row_data(row["row"], row)
                handler_config.save_workbook()
                return
        row = {"key": key, "value": value}
        handler_config.write_row_data(handler_config.get_max_row_number() + 1, row)
        handler_config.save_workbook()


def get_config_value(key):
    with ExcelSheetHandler(CACHE_EXCEL_NAME, CONFIG_SHEET_NAME) as handler_config:
        data = handler_config.get_all_data()
        for row in data:
            if row["key"] == key:
                return row["value"]
    return None


def create_all_path_sheet_header():
    if create_excel_sheet(CACHE_EXCEL_NAME, ALL_PATH_SHEET_NAME):
        with ExcelSheetHandler(CACHE_EXCEL_NAME, ALL_PATH_SHEET_NAME) as handler_all_path:
            handler_all_path.insert_column_header(1, "cs", "String", "path", "路径", 60)
            handler_all_path.insert_column_header(2, "cs", "String", "sheet_name", "页签名称,自动生成", 20)
            handler_all_path.insert_column_header(3, "cs", "Boolean", "includeSubDir", "是否包含子文件夹", 20)
            handler_all_path.insert_column_header(4, "cs", "String", "desc", "描述", 60)
            handler_all_path.save_workbook()


def set_path_data(path, sheet_name, include_subdirs, desc):
    with ExcelSheetHandler(CACHE_EXCEL_NAME, ALL_PATH_SHEET_NAME) as handler_all_path:
        data = handler_all_path.get_all_data()
        for row in data:
            if row["path"] == path:
                row["includeSubDir"] = include_subdirs
                row["desc"] = desc
                handler_all_path.write_row_data(row["row"], row)
                handler_all_path.save_workbook()
                return

        row = {"path": path, "includeSubDir": include_subdirs, "desc": desc, "sheet_name": sheet_name}
        handler_all_path.write_row_data(handler_all_path.get_max_row_number() + 1, row)
        handler_all_path.save_workbook()


def get_path_data(path):
    with ExcelSheetHandler(CACHE_EXCEL_NAME, ALL_PATH_SHEET_NAME) as handler_all_path:
        data = handler_all_path.get_all_data()
        for row in data:
            if row["path"] == path:
                return row
    return None


def get_all_path_data():
    with ExcelSheetHandler(CACHE_EXCEL_NAME, ALL_PATH_SHEET_NAME) as handler_all_path:
        return handler_all_path.get_all_data()


def get_path_sheet_name(path):
    with ExcelSheetHandler(CACHE_EXCEL_NAME, ALL_PATH_SHEET_NAME) as handler_all_path:
        data = handler_all_path.get_all_data()
        for row in data:
            if row["path"] == path:
                return row["sheet_name"]
    return None


def create_sheet_header(excel_name, sheet_name):
    if create_excel_sheet(excel_name, sheet_name):
        with ExcelSheetHandler(excel_name, sheet_name) as handler_sheet:
            handler_sheet.insert_column_header(1, "cs", "String", "name", "Excel文件名", 50)
            handler_sheet.insert_column_header(2, "cs", "String", "lastModified", "文件修改时间", 50)
            handler_sheet.insert_column_header(3, "cs", "String", "sheets", "页签列表", 50)
            handler_sheet.save_workbook()


def compute_cache_data():
    use_path = get_config_value("usePath")

    # 所有路径
    create_all_path_sheet_header()

    sheet_name = get_path_sheet_name(use_path)

    # 创建缓存文件
    create_sheet_header(CACHE_EXCEL_NAME, sheet_name)
    # 计算缓存数据
    handler = ExcelSheetHandler(CACHE_EXCEL_NAME, sheet_name)
    data = handler.get_all_data()
    # 遍历data转为字典key为文件名,value为行数据
    excel_map = {}
    for row in data:
        excel_map[row["name"]] = row

    # 统计逻辑耗时
    start_time = time.time()
    # 统计数量
    count = 0
    names = get_current_file_names(use_path, ".xlsx")
    for fullname in names:
        name = get_filename_from_path(fullname)
        last_modified = os.path.getmtime(fullname)
        # 转为时间格式
        last_modified = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(last_modified))
        # data中获取对应的数据
        if name in excel_map:
            row = excel_map[name]
            if row["lastModified"] == last_modified:
                continue
            else:
                # 获取页签列表，更新页签列表和修改时间
                sheet_names = get_excel_sheet_names(fullname)
                sheet_names = filter_sheet_names(sheet_names)
                row["lastModified"] = last_modified
                row["sheets"] = sheet_names
                handler.write_row_data(row["row"], row)
                logging.info("更新处理文件：" + name)
        else:
            # 新增数据
            sheet_names = get_excel_sheet_names(fullname)
            # 转换sheet_names为字符串,每个sheet_name之间用excel单元格的换行符分隔
            sheet_names = filter_sheet_names(sheet_names)
            row = {"cs": "cs", "name": name, "lastModified": last_modified, "sheets": sheet_names}
            handler.write_row_data(handler.get_max_row_number() + 1, row)
            logging.info("新增处理文件：" + name)
    # 保存
    handler.save_workbook()
    # 计算耗时
    t2 = time.time()
    print(f"计算缓存数据耗时：{t2 - start_time}秒")
    logging.info(f"计算缓存数据耗时：{t2 - start_time}秒")


def get_all_sheet_names():
    use_path = get_config_value("usePath")
    sheet_name = get_path_sheet_name(use_path)
    with ExcelSheetHandler(CACHE_EXCEL_NAME, sheet_name) as handlerRead:
        data = handlerRead.get_all_data()
        # 遍历 data 转为字典，key 为文件名，value 为行数据
        sheet_names = []
        for row in data:
            sheets = row["sheets"]
            if sheets:
                for sheet_name in sheets.split('\n'):
                    sheet_names.append({"name": row["name"], "sheet_name": sheet_name})
        return sheet_names


# 过滤不包含|的页签名，并转为\n分隔的字符串
def filter_sheet_names(sheet_names):
    # return "\n".join([sheet_name for sheet_name in sheet_names if "|" in sheet_name])
    return "\n".join([sheet_name for sheet_name in sheet_names])


# 判断当前缓存是否存在页签
def exist_sheet(sheet_name):
    names = get_excel_sheet_names(CACHE_EXCEL_NAME)
    return sheet_name in names


if __name__ == '__main__':
    excel_name = "cache.xlsx"
    sheet_name = "缓存文件列表"
    # 计算缓存数据
    handler = ExcelSheetHandler(excel_name, sheet_name)
    handler.clear_data_rows()
    handler.save_workbook()
