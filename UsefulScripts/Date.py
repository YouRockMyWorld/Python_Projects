import datetime


def get_date_list(start_date = None, end_date = None, step = 1, format = '%Y-%m-%d'):
    '''
    得到从起始日期到结束日期之间的日期字符串序列,若start_date,end_date非None,需要指定format
    :param start_date: 开始日期，缺省则默认为'2019-01-01'
    :param end_date: 结束日期，缺省则默认为今天日期'xxxx-xx-xx'
    :param step: 步长，缺省则默认则为1
    :param format: 格式，缺省则默认为'%Y-%m-%d'
    :return: 日期列表
    '''
    if start_date is None:
        start_date = datetime.datetime.strptime('2019-01-01', '%Y-%m-%d').strftime(format)
    if end_date is None:
        end_date = datetime.datetime.now().strftime(format)

    start = datetime.datetime.strptime(start_date, format)
    end = datetime.datetime.strptime(end_date, format)

    datelist = []
    while start <= end:
        datelist.append(start.strftime(format))
        start+=datetime.timedelta(days=step)

    return datelist


