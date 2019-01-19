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
                self.__ProcessXlsx()
            elif self.writePath[-1] == 's':
                self.__ProcessXls()
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
    def ToUpper(self, s, sheet_name):
        if s and str(s)[0].isalpha() and sheet_name == '地表沉降':
            return s.upper()[1:]
        else:
            return s

    # 原始ToUpper函数，只将所有点号字母改为大写
    # def ToUpper(self, s):
    #     if s and str(s)[0].isalpha():
    #         return s.upper()
    #     else:
    #         return s


    def __ProcessXls(self):
        out_info = ''
        read_excel_book = xlrd.open_workbook(self.readPath)
        write_excel_book = xlwt.Workbook()
        # write_excel_book = xlsxwriter.Workbook(write_excel_book_path)
        x = 1
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
                    finallist.append(list(map(lambda s: self.ToUpper(s, sheetconf['WriteSheetName']), l)))
                    # finallist.append(list(map(lambda s: self.ToUpper(s), l)))


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

            except:
                continue

        out_info += '*' * 100 + '\n'
        self.signalOut.emit('*' * 100 + '\n')
        write_excel_book.save(self.writePath)
        return out_info


    def __ProcessXlsx(self):
        out_info = ''
        read_excel_book = xlrd.open_workbook(self.readPath)
        write_excel_book = xlsxwriter.Workbook(self.writePath)
        x = 1
        for sheetconf in self.conf:
            try:
                print('第 %d 个sheet，名字为：%s' % (x, sheetconf['WriteSheetName']))
                self.signalOut.emit('第 %d 个sheet，名字为：%s   ' % (x, sheetconf['WriteSheetName']) + '\n')
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
                    finallist.append(list(map(lambda s: self.ToUpper(s, sheetconf['WriteSheetName']), l)))
                    # finallist.append(list(map(lambda s: self.ToUpper(s), l)))


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

            except:
                continue

        out_info += '*' * 100 + '\n'
        self.signalOut.emit('*' * 100 + '\n')
        write_excel_book.close()
        return out_info
