from File import *
from Date import *
import re, os

if __name__ == '__main__':
    files = get_all_files(r'C:\Users\Administrator\Desktop\集成村站转换\xls', '*.xls')
    files.sort(key=lambda x: int(re.findall(r'\d+', x)[0]))
    date = get_date_list(start_date='_2018-06-23', end_date='_2018-12-27', format='_%Y-%m-%d')
    print(date)
    print(len(date))
    print(len(files[0:-3]))
    i = 0
    for file in files[0:-3]:
        index = file.find('_')
        newname = file[0:index].strip() + date[i] + '.xls'
        i+=1
        print(file)
        print(newname)
        os.rename(file, newname)