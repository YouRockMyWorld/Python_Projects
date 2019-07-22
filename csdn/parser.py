from bs4 import BeautifulSoup
import random
import CONF
import requests
import time


def get_page(user_name):
    try:
        url = r'https://blog.csdn.net/{0}'.format(user_name)
        print(url)
        headers = CONF.headers
        headers['User-Agent'] = random.choice(CONF.user_agent_list)
        res = requests.get(url, headers=headers)
        print(res.status_code)
        if res.status_code == 200:
            return res.text
        else:
            return None
    except:
        return None

def get_user_list(html):
    href_list = set()
    soup = BeautifulSoup(html, 'lxml')
    items = soup.select('div.article-list a')
    #items = soup.find_all(class_='article-list')

    for item in items:
        href_list.add(item['href'])

    return list(href_list)

def access_url(article_list):
    for i in range(CONF.access_count):
        try:
            headers = CONF.headers
            headers['User-Agent'] = random.choice(CONF.user_agent_list)
            url = random.choice(article_list)
            res = requests.get(url, headers=headers)
            if res.status_code == 200:
                print('正在循环第{0}次访问，url={1}'.format(i, url))
            else:
                print('第{0}次访问失败，url={1}'.format(i, url))
            time.sleep(random.randint(CONF.interval[0], CONF.interval[1]))
        except:
            continue




if __name__ == '__main__':
    html = get_page(CONF.user_name)
    article_list = []
    if html:
        article_list = get_user_list(html)

        for item in article_list:
            if CONF.user_name not in item:
                article_list.remove(item)
    access_url(article_list)




