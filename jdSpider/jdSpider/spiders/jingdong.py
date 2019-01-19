# -*- coding: utf-8 -*-
from scrapy import Request, Spider
from jdSpider.items import JdspiderItem


class JingdongSpider(Spider):
    name = 'jingdong'
    allowed_domains = ['www.jd.com']
    #start_urls = ['http://www.jd.com/']
    base_url = 'https://list.jd.com/list.html?cat=670,671,672'

    def start_requests(self):
        for page in range(1, self.settings.get('MAX_PAGE') + 1):
            yield Request(url = self.base_url, callback = self.parse, meta = {'page':page}, dont_filter = True)

    def parse(self, response):
        products = response.css('ul.gl-warp.clearfix li div.gl-i-wrap.j-sku-item')
        for product in products:
            item = JdspiderItem()
            item['title'] = product.css('div.p-name em::text').extract_first('').strip()
            item['price'] = product.css('div.p-price i::text').extract_first('').strip()
            item['shop'] = product.css('div.p-shop span a::text').extract_first('').strip()
            yield item

