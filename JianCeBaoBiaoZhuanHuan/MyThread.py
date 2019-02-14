from PyQt5 import QtCore
import xlrd, xlwt
import xlsxwriter
import time


class ExcelThread(QtCore.QThread):
    signalOut = QtCore.pyqtSignal(str)

    def __init__(self, parent=None):
        super(ExcelThread, self).__init__()
        self.readPath = ''
        self.writePath = ''
        self.conf = None

    def run(self):
        try:
            if self.writePath[-1] == 'x':
                self.ProcessXlsx(self.readPath, self.writePath)
            elif self.writePath[-1] == 's':
                self.ProcessXls(self.readPath, self.writePath)
        except Exception as e:
            self.signalOut.emit(str(e))



    # 将数据墙顶水平位移数据除以1000，不需要的话使用注释中的 ToUpper 函数，同时记得将两个调用改掉
    # def ToUpper(self, s, sheet_name):
    #     if s and str(s)[0].isalpha():
    #         return s.upper()
    #     elif s and sheet_name == '墙顶变形_墙顶水平位移':
    #         return s / 1000.0
    #     else:
    #         return s


    # 将地表沉降点号去掉“D”只留“B1-2”，不需要的话使用注释中的 ToUpper 函数，同时记得将两个调用改掉
    # def ToUpper(self, s, sheet_name):
    #     if s and str(s)[0].isalpha() and sheet_name == '地表沉降':
    #         return s.upper()[1:]
    #     else:
    #         return s

    # 原始ToUpper函数，只将所有点号字母改为大写
    def ToUpper(self, s):
        if s and str(s)[0].isalpha():
            return s.upper()
        else:
            return s


    def delete_null_value_in_list(self, l):
        ss = set(map(lambda x: '' in x, l))
        if len(ss) == 1 and ss.pop() == True:
            min_len = min(list(map(len, l)))
            for i in range(min_len - 1, -1 , -1):
                s = set(map(lambda x: x[i], l))
                if len(s) == 1 and s.pop() == '':
                    for li in l:
                        li.pop(i)



    def ProcessXls(self, read_path, write_path):
        out_info = ''
        read_excel_book = xlrd.open_workbook(read_path)
        sheet_list = read_excel_book.sheet_names()
        write_excel_book = xlwt.Workbook()
        # write_excel_book = xlsxwriter.Workbook(write_excel_book_path)
        x = 1
        print('正在将："' + read_path + '" 写入 "' + write_path + '"')
        self.signalOut.emit('正在将："' + read_path + '" 写入 "' + write_path + '"\n')
        for sheetconf in self.conf:
            try:
                print('第 %d 个sheet，名字为：%s' % (x, sheetconf['WriteSheetName']))
                self.signalOut.emit('第 %d 个sheet，名字为：%s   ' % (x, sheetconf['WriteSheetName']) + '\n')
                out_info += '第 %d 个sheet，名字为：%s   ' % (x, sheetconf['WriteSheetName'])
                write_sheet = write_excel_book.add_sheet(sheetconf['WriteSheetName'])
                # write_sheet = write_excel_book.add_worksheet(sheetconf['WriteSheetName'])
                added_header_name = []
                valuelist = []
                for item in sheetconf['Read']:
                    if item['ReadSheetName'] not in sheet_list:
                        continue
                    read_sheet = read_excel_book.sheet_by_name(item['ReadSheetName'])
                    for dataitem in item['Data']:
                        value = []
                        # valuelist.append({dataitem['WriteHeaderName']:[]})
                        for readdataitem in dataitem['ReadData']:
                            value.extend(read_sheet.col_values(readdataitem['Column'] - 1)[
                                         readdataitem['StartRow'] - 1: readdataitem['EndRow']])

                        valuelist.append({dataitem['WriteHeaderName']: value})

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
                    # finallist.append(list(map(lambda s: self.ToUpper(s, sheetconf['WriteSheetName']), l)))
                    finallist.append(list(map(lambda s: self.ToUpper(s), l)))

                # 验证是否存在有空值,若有则删除
                self.delete_null_value_in_list(finallist)

                for item in header:
                    write_sheet.write(0, header.index(item), item)
                    row_index = 1
                    col_index = header.index(item)
                    for col_data in finallist[col_index]:
                        write_sheet.write(row_index, col_index, col_data)
                        row_index += 1
                    # write_sheet.write_column(1, header.index(item), finallist[header.index(item)])

                print('header:')
                out_info += 'header: '
                self.signalOut.emit('header: ')
                print(header)
                out_info += str(header) + '   '
                self.signalOut.emit(str(header) + '   \n')
                print('data:')
                out_info += 'data: '
                self.signalOut.emit('data: ')
                print(finallist)
                out_info += str(finallist) + '   '
                self.signalOut.emit(str(finallist) + '   \n')
                print('\n')
                out_info += '\n'
                self.signalOut.emit('\n')
                x += 1
                time.sleep(0.1)

            except Exception as e:
                print(str(e) + '\n')
                continue

        out_info += '*' * 100 + '\n'
        self.signalOut.emit('*' * 100 + '\n')
        write_excel_book.save(write_path)
        return out_info


    def ProcessXlsx(self, read_path, write_path):
        out_info = ''
        read_excel_book = xlrd.open_workbook(read_path)
        sheet_list = read_excel_book.sheet_names()
        write_excel_book = xlsxwriter.Workbook(write_path)
        x = 1
        print('正在将："' + read_path + '" 写入 "' + write_path + '"')
        self.signalOut.emit('正在将："' + read_path + '" 写入 "' + write_path + '"\n')
        for sheetconf in self.conf:
            try:
                print('第 %d 个sheet，名字为：%s' % (x, sheetconf['WriteSheetName']))
                self.signalOut.emit('第 %d 个sheet，名字为：%s   ' % (x, sheetconf['WriteSheetName']) + '\n')
                out_info += '第 %d 个sheet，名字为：%s   ' % (x, sheetconf['WriteSheetName'])
                write_sheet = write_excel_book.add_worksheet(sheetconf['WriteSheetName'])
                added_header_name = []
                valuelist = []
                for item in sheetconf['Read']:
                    if item['ReadSheetName'] not in sheet_list:
                        continue
                    read_sheet = read_excel_book.sheet_by_name(item['ReadSheetName'])
                    for dataitem in item['Data']:
                        value = []
                        # valuelist.append({dataitem['WriteHeaderName']:[]})
                        for readdataitem in dataitem['ReadData']:
                            value.extend(read_sheet.col_values(readdataitem['Column'] - 1)[
                                         readdataitem['StartRow'] - 1: readdataitem['EndRow']])

                        valuelist.append({dataitem['WriteHeaderName']: value})

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
                    # finallist.append(list(map(lambda s: self.ToUpper(s, sheetconf['WriteSheetName']), l)))
                    finallist.append(list(map(lambda s: self.ToUpper(s), l)))

                # 验证是否存在有空值,若有则删除
                self.delete_null_value_in_list(finallist)

                for item in header:
                    write_sheet.write(0, header.index(item), item)
                    write_sheet.write_column(1, header.index(item), finallist[header.index(item)])


                print('header:')
                out_info += 'header: '
                self.signalOut.emit('header: ')
                print(header)
                out_info += str(header) + '   '
                self.signalOut.emit(str(header) + '   \n')
                print('data:')
                out_info += 'data: '
                self.signalOut.emit('data: ')
                print(finallist)
                out_info += str(finallist) + '   '
                self.signalOut.emit(str(finallist) + '   \n')
                print('\n')
                out_info += '\n'
                self.signalOut.emit('\n')
                x += 1
                time.sleep(0.1)

            except Exception as e:
                print(str(e) + '\n')
                continue

        out_info += '*' * 100 + '\n'
        self.signalOut.emit('*' * 100 + '\n')
        write_excel_book.close()
        return out_info


class MultiExcelThread(ExcelThread):
    def __init__(self):
        super(MultiExcelThread, self).__init__()
        self.readPath_list = []

    def run(self):
        try:
            for readpath in self.readPath_list:
                if readpath[-1] == 'x':
                    writepath = readpath.replace('.xlsx', '_转换.xlsx')
                    self.ProcessXlsx(readpath, writepath)
                elif readpath[-1] == 's':
                    writepath = readpath.replace('.xls', '_转换.xls')
                    self.ProcessXls(readpath, writepath)
        except Exception as e:
            self.signalOut.emit(str(e))