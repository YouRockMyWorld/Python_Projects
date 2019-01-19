import xlrd, xlwt
import json
import xlsxwriter


def ProcessExcelData(conf, read_excel_book, write_excel_book_name):
    read_sheets = read_excel_book.sheets()
    added_sheet_names = []

    # 对于每一个配置文件中的object处理
    for sheetconf in conf:
        write_excel_book = xlsxwriter.Workbook(write_excel_book_name)
        if sheetconf['WriteSheetName'] in added_sheet_names:
            write_sheet = write_excel_book.get_worksheet_by_name(sheetconf['WriteSheetName'])
            read_sheet = read_excel_book.sheet_by_name(sheetconf['ReadSheetName'])
            row_values = write_sheet.row(0)
            print(row_values)

            for dataitem in sheetconf['Data']:
                row_index = write_sheet.nrows
                col_index = row_values.index(sheetconf['WriteSheetName'])
                # 根据配置读取原始数据到列表
                valuelist = []
                for readdataitem in dataitem['ReadData']:
                    valuelist.extend(read_sheet.col_values(readdataitem['Column'] - 1)[
                                     readdataitem['StartRow'] - 1: readdataitem['EndRow']])

                # 将列表数据写入新sheet列
                for value in valuelist:
                        write_sheet.write(row_index, col_index, value)
                        row_index += 1

        else:
            # 对每一个sheet名字和配置文件中object['ReadSheetName']做匹配，匹配到就按照配置项目处理读取该sheet并写入新文件
            for read_sheet in read_sheets:
                #print(read_sheet.name + sheetconf['ReadSheetName'])
                if (read_sheet.name == sheetconf['ReadSheetName']):
                    write_sheet = write_excel_book.add_worksheet(sheetconf['WriteSheetName'])
                    added_sheet_names.append(sheetconf['WriteSheetName'])
                    col_index = 0
                    for dataitem in sheetconf['Data']:
                        row_index = 0
                        write_sheet.write(row_index, col_index, dataitem['WriteHeaderName'])
                        row_index += 1

                        # 根据配置读取原始数据到列表
                        valuelist = []
                        for readdataitem in dataitem['ReadData']:
                            valuelist.extend(read_sheet.col_values(readdataitem['Column'] - 1)[
                                             readdataitem['StartRow'] - 1: readdataitem['EndRow']])

                        # 将列表数据写入新sheet列
                        for value in valuelist:
                            write_sheet.write(row_index, col_index, value)
                            row_index += 1

                        col_index += 1
        write_excel_book.close()
    print('out')


def ToUpper(s):
    if s and str(s)[0].isalpha():
        return s.upper()
    else:
        return s

#将数据墙顶水平位移数据除以1000，不需要的话使用注释中的 ToUpper 函数，同时记得将两个调用改掉
# def ToUpper(s, sheet_name):
#     if s and str(s)[0].isalpha():
#         return s.upper()
#     elif s and sheet_name == '墙顶变形_墙顶水平位移':
#         return s / 1000.0
#     else:
#         return s


#写xls
def ProcessXls(conf, read_excel_book_path, write_excel_book_path):
    out_info = ''
    read_excel_book = xlrd.open_workbook(read_excel_book_path)
    write_excel_book = xlwt.Workbook()
    #write_excel_book = xlsxwriter.Workbook(write_excel_book_path)
    x = 1
    for sheetconf in conf:
        print('第 %d 个sheet，名字为：%s' % (x, sheetconf['WriteSheetName']))
        out_info += '第 %d 个sheet，名字为：%s   ' % (x, sheetconf['WriteSheetName'])
        write_sheet = write_excel_book.add_sheet(sheetconf['WriteSheetName'])
        #write_sheet = write_excel_book.add_worksheet(sheetconf['WriteSheetName'])
        added_header_name = []
        valuelist = []
        for item in sheetconf['Read']:
            read_sheet = read_excel_book.sheet_by_name(item['ReadSheetName'])
            for dataitem in item['Data']:
                value = []
                # valuelist.append({dataitem['WriteHeaderName']:[]})
                for readdataitem in dataitem['ReadData']:
                    value.extend(read_sheet.col_values(readdataitem['Column'] - 1)[readdataitem['StartRow'] - 1: readdataitem['EndRow']])

                valuelist.append({dataitem['WriteHeaderName'] : value})

        header = []
        for item in valuelist:
            if list(item.keys())[0] not in header:
                header.append(list(item.keys())[0])

        finallist = []
        for item in header:
            l = []
            for i in valuelist:
                if list(i.keys())[0] == item:
                    l.extend(list(i.values())[0])
            #finallist.append(list(map(lambda s:ToUpper(s, sheetconf['WriteSheetName']), l)))
            finallist.append(list(map(lambda s:ToUpper(s), l)))


        for item in header:
            write_sheet.write(0,header.index(item),item)
            row_index = 1
            col_index = header.index(item)
            for col_data in finallist[col_index]:
                write_sheet.write(row_index, col_index, col_data)
                row_index+=1
            #write_sheet.write_column(1, header.index(item), finallist[header.index(item)])

        print('header:')
        out_info += 'header: '
        print(header)
        out_info += str(header) + '   '
        print('data:')
        out_info += 'data: '
        print(finallist)
        out_info += str(finallist) + '   '
        print('\n')
        out_info += '\n'
        x += 1

    out_info += '*'*100 + '\n'
    write_excel_book.save(write_excel_book_path)
    return out_info


#写xlsx
def ProcessXlsx(conf, read_excel_book_path, write_excel_book_path):

    out_info = ''
    read_excel_book = xlrd.open_workbook(read_excel_book_path)
    write_excel_book = xlsxwriter.Workbook(write_excel_book_path)
    x = 1
    for sheetconf in conf:
        print('第 %d 个sheet，名字为：%s' % (x, sheetconf['WriteSheetName']))
        out_info += '第 %d 个sheet，名字为：%s   ' % (x, sheetconf['WriteSheetName'])
        write_sheet = write_excel_book.add_worksheet(sheetconf['WriteSheetName'])
        added_header_name = []
        valuelist = []
        for item in sheetconf['Read']:
            read_sheet = read_excel_book.sheet_by_name(item['ReadSheetName'])
            for dataitem in item['Data']:
                value = []
                # valuelist.append({dataitem['WriteHeaderName']:[]})
                for readdataitem in dataitem['ReadData']:
                    value.extend(read_sheet.col_values(readdataitem['Column'] - 1)[readdataitem['StartRow'] - 1: readdataitem['EndRow']])

                valuelist.append({dataitem['WriteHeaderName'] : value})

        header = []
        for item in valuelist:
            if list(item.keys())[0] not in header:
                header.append(list(item.keys())[0])

        finallist = []
        for item in header:
            l = []
            for i in valuelist:
                if list(i.keys())[0] == item:
                    l.extend(list(i.values())[0])
            # finallist.append(list(map(lambda s:ToUpper(s, sheetconf['WriteSheetName']), l)))
            finallist.append(list(map(lambda s:ToUpper(s), l)))


        for item in header:
            write_sheet.write(0,header.index(item),item)
            write_sheet.write_column(1, header.index(item), finallist[header.index(item)])


        print('header:')
        out_info += 'header: '
        print(header)
        out_info += str(header) + '   '
        print('data:')
        out_info += 'data: '
        print(finallist)
        out_info += str(finallist) + '   '
        print('\n')
        out_info += '\n'
        x += 1

    out_info += '*' * 100 + '\n'
    write_excel_book.close()
    return out_info



# if __name__ == '__main__':
#     with open('conf_map_platform_new.json', encoding='utf-8') as f:
#         jsonstr = f.read()
#     dataconf = json.loads(jsonstr)
#
#
#     # ProcessExcelData(dataconf, read_excel_book, write_excel_book)
#     Process(dataconf, r'C:\Users\Administrator\Desktop\test\集成村站施工监测日报第140期.xlsx', 'platform_out_newt.xlsx')
