# -*- coding:utf-8 -*-
from collections import OrderedDict
import xlrd


def readExcel(file, table_name=None):
    """
    获取excel表内数据
    :param file: 文件名
    :param table_name: 表名
    :return:
    """
    wb = xlrd.open_workbook(filename=file)

    if table_name:
        sheet = wb.sheet_by_name(table_name)  # 通过名字获取表格
    else:
        sheet = wb.sheet_by_index(0)

    sheet_rows = sheet.nrows  # 总行数
    sheet_cols = sheet.ncols  # 总列数

    table_head = [sheet.cell_value(0, i) for i in range(sheet_cols)]

    table_body = []

    for x in range(1, sheet_rows):
        tmp = [sheet.cell_value(x, i) for i in range(sheet_cols)]
        table_body.append(tmp)
    result = {
        'table_head': table_head,
        'table_body': table_body
    }
    return result


def get_setting(config):
    file = "配置表.xlsx"
    result = None
    if config == "商品":
        result = {}
        data = readExcel(file, config)
        for i in data['table_body']:
            try:
                length = float(i[1]) if i[1] else 0
                width = float(i[2]) if i[2] else 0
                height = float(i[3]) if i[3] else 0

                result[i[0]] = {
                    "length": length,
                    "width": width,
                    "height": height,
                    "volume": length * width * height
                }
            except Exception as err:
                print(i, err)

    if config == "箱型":
        result = []
        data = readExcel(file, config)
        for i in data["table_body"]:
            try:
                name = i[0] + "-" + i[1]
                length = float(i[2]) if i[2] else 0
                width = float(i[3]) if i[3] else 0
                height = float(i[4]) if i[4] else 0
                volume = length * width * height
                tmp = {
                    "name": name,
                    "length": length,
                    "width": width,
                    "height": height,
                    "volume": volume
                }
                result.append(tmp)
            except Exception as err:
                print(i, err)

        for i in range(len(result)):
            min_idx = i
            for j in range(i + 1, len(result)):
                if result[min_idx]['volume'] > result[j]['volume']:
                    min_idx = j
            result[i], result[min_idx] = result[min_idx], result[i]

    return result


if __name__ == "__main__":
    print(get_setting("箱型"))
