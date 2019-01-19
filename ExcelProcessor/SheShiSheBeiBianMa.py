import xlrd
import xlsxwriter

def process_sheet(sheet):
    '''
    从某张sheet中取出数据
    :param sheet: 表
    :return: 返回列表数据
    '''
    data = []
    for r in range(sheet.nrows):
        row = []
        for c in range(6):
            value = str(sheet.cell_value(r, c))
            if value.strip():
                row.append(sheet.cell_value(r, c))
            elif c <= 2:
                row.append('00')
            else:
                row.append('000')
        data.append(row)
    # for i in data:
    #     for j in i:
    #         print(j)
    #     print('-----')
    return data


def write_new_workbook(workbook, sheetname, data, *format):
    '''
    往excel中写数据
    :param workbook:
    :param sheetname:
    :param data:
    :param format:
    :return:
    '''
    sheet = workbook.add_worksheet(sheetname)
    for r in range(len(data)):
        sheet.write_row(r, 0, data[r], *format)


if __name__ == '__main__':
    read_workbook = xlrd.open_workbook(r'C:\Users\Administrator\Desktop\新建文件夹\分类代码表.xlsx')
    write_workbook = xlsxwriter.Workbook(r'C:\Users\Administrator\Desktop\新建文件夹\out.xlsx')
    cell_format = write_workbook.add_format({
        'font_name': '微软雅黑',
        'font_size': 11,
        'border': 1,
        'align': 'center',
        'valign': 'vcenter'})
    for i in range(1, read_workbook.nsheets):
        sheet = read_workbook.sheet_by_index(i)
        data = process_sheet(sheet)
        write_new_workbook(write_workbook, sheet.name, data, cell_format)
    write_workbook.close()