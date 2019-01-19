from urllib.parse import urlencode
from urllib.request import urlretrieve
from requests.exceptions import RequestException
import requests
import time
import random
import json
import os
from utils import *


#请求url,获取url内容
def get_page(keyword, index):
    url = 'https://image.baidu.com/search/acjson?'
    params = {
        'tn': 'resultjson_com',
        'ipn': 'rj',
        'ct': '201326592',
        'is': '',
        'fp': 'result',
        'queryWord': keyword,
        'cl': '2',
        'lm': '-1',
        'ie': 'utf-8',
        'oe': 'utf-8',
        'adpicid': '',
        'st': '-1',
        'z': '',
        'ic': '0',
        'word': keyword,
        's': '',
        'se': '',
        'tab': '',
        'width': '',
        'height': '',
        'face': '0',
        'istype': '2',
        'qc': '',
        'nc': '1',
        'fr': '',
        'pn': index * 30,
        'rn': '30',
        'gsm': str(hex(index * 30))[2:],
        round(time.time() * 1000): ''
    }
    url += urlencode(params)
    # print(url)
    try:
        headers['User-Agent'] = random.choice(user_agent)
        response = requests.get(url, headers=headers)
        response.encoding = 'utf-8'
        if response.status_code == 200:
            return response.text
        return None
    except RequestException:
        print('"get_page"请求失败:' + url)
        return None


#解析url返回内容，返回字典封装的一个图片信息
def parse_page(content):
    data = json.loads(content)
    #print(data)
    if data:
        for item in data.get('data'):
            yield {
                'title':item.get('fromPageTitleEnc'),
                'hoverURL':item.get('hoverURL'),
                'middleURL':item.get('middleURL'),
                'thumbURL':item.get('thumbURL'),
                'objURL':decode_url(str(item.get('objURL'))) if item.get('objURL') else None,
                'type':item.get('type'),
            }


# def get_image(url):
#     try:
#         headers['User-Agent'] = random.choice(user_agent)
#         response = requests.get(url, headers=headers)
#         if response.status_code == 200:
#             return response.content
#         return None
#     except RequestException:
#         print('"get_image"请求失败:' + url)
#         return None


#下载图片，以网页中显示的标题作为文件名
def save_image(content, keyword):
    try:
        if content['title']:
            path = '%s/%s' % (os.getcwd(), keyword)
            if not os.path.exists(path):
                os.mkdir(path)
            urlretrieve(content['objURL'], os.path.join(path, get_legal_name(content['title'] + '.' + content['type'])))
            print('downloading ' + content['title'])
            #定义间隔时间，防止不使用代理封IP
            time.sleep(1)
    except:
        print('"save_image"失败')


#对文件名进行Windows下的合法化
def get_legal_name(file_name):
    result = re.findall(r'[\\/:*?"<>|\r\n]+', file_name)
    if result:
        for i in result:
            file_name = file_name.replace(i, '_')
    return file_name


#爬取一条记录相关所有信息
def get_one_class(keyword, page):
    for i in range(page):
        content = get_page(keyword, i)
        for item in parse_page(content):
            print(item)
            save_image(item, keyword)


#从文件中得到要爬虫的信息
def get_info(file_name):
    with open(file_name, 'r') as f:
        content = f.readlines()
    mlist = []
    for line in content:
        s = line.strip('\n').split(';')
        mlist.append({
            'keyword':s[0],
            'page':s[1]
        })
    return mlist


#主函数
def main(file_name):
    content = get_info(file_name)
    for item in content:
        get_one_class(item['keyword'], int(item['page']))



if __name__ == '__main__':
    main('./spider.txt')
