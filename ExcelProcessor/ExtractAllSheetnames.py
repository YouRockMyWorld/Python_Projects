import xlrd, xlsxwriter
import re, os
from PiLiangXieShuJu import get_all_files


def get_sheet_names(path):
    workbook = xlrd.open_workbook(path)
    return workbook.sheet_names()


def get_number_in_path(path):
    filename = os.path.split(path)[1]
    result = re.findall(r'\d+', filename)
    return result[0]


if __name__ == '__main__':
    files = get_all_files()
    workbook = xlsxwriter.Workbook(r'C:\Users\Administrator\Desktop\新建文件夹\result.xlsx')
    worksheet = workbook.add_worksheet('result')
    row = 0
    for filename in files:
        try:
            print('正在处理:', filename)
            data = get_sheet_names(filename)
            data.insert(0, get_number_in_path(filename))
            worksheet.write_row(row, 0, data)
            row += 1
        except:
            continue
    workbook.close()