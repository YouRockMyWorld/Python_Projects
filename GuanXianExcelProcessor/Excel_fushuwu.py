import xlrd,xlwt


def Process_Sheet(sheet):
    result_list = []
    rows_list = []
    lianjiedian_col_list = []
    processed_set = set([])
    for i in range(3,sheet.nrows):
        row = sheet.row_values(i)
        rows_list.append(row)
        lianjiedian_col_list.append(row[1])


    for row in rows_list:
        #aim_index = lianjiedian_col_list.index(row[0])
        #result = [row[0],row[1],row[4],row[5],row[8]*1000,rows_list[aim_index][4],rows_list[aim_index][5],rows_list[aim_index][8]*1000,row[9],row[10]]
        if row[3] != '':
            result = []
            if 'X' in str(row[9]):
                slist = row[9].split('X')
                if float(slist[0])>=300 or float(slist[1])>=300:
                    result = [row[0], row[1], row[3], row[5] * 1000, row[4] * 1000, row[7] * 1000, row[9], row[10]]
            else:
                if float(row[9]) >=300:
                    result = [row[0], row[1], row[3], row[5] * 1000, row[4] * 1000, row[7] * 1000, row[9], row[10]]

            pointstr = row[0]
            if (pointstr not in processed_set) and result:
                result_list.append(result)
                processed_set.add(pointstr)

    return result_list




if __name__ == '__main__':
    exceldata = xlrd.open_workbook(r"C:\Users\Administrator\Desktop\最新.xlsx")
    sheets = exceldata.sheets()
    for sheet in sheets:
        with open(sheet.name + '.txt', 'w', encoding='utf-8') as f:
            for row in Process_Sheet(sheet):
                s = ''
                for i in row:
                    s += str(i)+'\t'
                f.writelines([s,'\n'])

        # print(sheet.name)
        # for i in Process_Sheet(sheet):
        #     print(i)


