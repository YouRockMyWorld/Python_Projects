# -*- coding: utf-8 -*-

# Define your item pipelines here
#
# Don't forget to add your pipeline to the ITEM_PIPELINES setting
# See: https://doc.scrapy.org/en/latest/topics/item-pipeline.html
import os

class JdspiderPipeline(object):
    def __init__(self,filename):
        self.filename = filename
        self.filepath = os.path.join(os.getcwd(), self.filename)

    @classmethod
    def from_crawler(cls,crawler):
        return cls(filename=crawler.settings.get('FILENAME'))


    def open_spider(self,spider):
        self.file = open(self.filepath, 'a', encoding='utf-8')

    def process_item(self, item, spider):
        self.file.write(str(item) + '\n' + '-'*100 + '\n')
        return item

    def close_spider(self,spider):
        self.file.close()
