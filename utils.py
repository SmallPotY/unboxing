# -*- coding:utf-8 -*-
import xlrd
import xlwt
import datetime
import os
import sys

application_path = ""
if getattr(sys, 'frozen', False):
    application_path = os.path.dirname(sys.executable)
elif __file__:
    application_path = os.path.dirname(__file__)


def writeExcel(content):
    """
    写入excel
    :param content:  内容列表，包含表头
    :return: 返回文件路径
    """
    workbook = xlwt.Workbook(encoding='utf-8')
    sheet = workbook.add_sheet("sheet1")

    # 遍历内容写入
    for r_index, row in enumerate(content):
        for c_index, column in enumerate(row):
            sheet.write(r_index, c_index, content[r_index][c_index])

    fileName = "箱型计算_" + datetime.datetime.now().strftime("%Y%m%d%H%M%S") + '.xls'
    workbook.save(fileName)  # 保存
    path = os.path.join(application_path, fileName)
    return path


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
                    "volume": volume,
                    "diagonal": diagonal(length, width, height)
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


def diagonal(*length_width_height):
    """
    求对角线长度
    :return:
    """
    diagonal_list_tmp = []
    for a_index, a in enumerate(length_width_height):
        for b_index, b in enumerate(length_width_height):
            if a_index != b_index:
                diagonal_list_tmp.append([a, b])
    diagonal_list = set()
    for i in diagonal_list_tmp:
        v = (i[0] ** 2 + i[1] ** 2) ** (1 / 2)
        diagonal_list.add(v)

    return list(diagonal_list)


if __name__ == "__main__":
    t = get_setting("箱型")
    print(t)
