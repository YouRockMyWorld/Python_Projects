# -*- coding: utf-8 -*-
import scrapy
from urllib.parse import urlencode

#定义爬虫
class RenrenSpider(scrapy.Spider):
    name = 'renren'
    allowed_domains = ['www.renren.com']
    #输入账号密码
    email = ''
    password = ''


#定义爬虫请求
    def start_requests(self):
        url = 'http://www.renren.com/PLogin.do'
        #url = 'http://www.renren.com/ajaxLogin/login?1=1&uniqueTimestamp=2018711536567'
        data = {
                'email': self.email,
                'password':self.password,
            }

        yield scrapy.FormRequest(
            url = url,
            formdata = data,
        )
#定义解析函数
    def parse(self, response):
        if response.status == 200:
            print('=' * 70)
            print('欢迎您')
            print(response.css('head title::text').extract_first())
            print('='*70)
            with open('./test.html', 'w', encoding='utf=8') as f:
                f.write(response.body.decode('utf-8'))
