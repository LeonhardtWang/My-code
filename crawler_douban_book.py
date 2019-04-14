#!/usr/bin/env python
# -*- encoding: utf-8 -*-
'''
@File    :   crawler_douban_book.py
@Time    :   2019/04/12 22:54:35
@Author  :   LeoWang 
@Version :   1.0
@Contact :   953744841@qq.com
@License :   (C)Copyright 2017-2018, Liugroup-NLPR-CASIA
@Desc    :   None
'''

# here put the import lib
from selenium import webdriver
from lxml import etree
import time
import re
import pandas as pd

class Crawler:
    '''
    爬虫豆瓣上书籍信息,并按评分降序排序
    '''
    def __init__(self, key_word):
        self.driver = webdriver.Firefox()
        self.url = 'https://book.douban.com/'
        self.key_word = key_word
        self.res_list = []

    def get_ele(self, select_str, is_css=True):
        try:
            if is_css:
                ele = self.driver.find_element_by_css_selector(select_str)
            else:
                ele = self.driver.find_element_by_xpath(select_str)
            return ele
        except:
            return None

    def click(self, ele):
        self.driver.execute_script('arguments[0].click();', ele)
        time.sleep(3)

    def get_path(self):
        local_tiem = time.strftime("%Y.%m.%d.%H.%M.%S", time.localtime())
        return local_tiem + '_' + self.key_word.replace('"', '') + '.csv'

    def search(self):
        self.driver.get(self.url)
        input_val = self.get_ele('input#inp-query')
        input_val.send_keys(self.key_word)
        button = self.get_ele('input[type="submit"][value="搜索"]')
        self.click(button)

    def get_items(self):
        page_source = self.driver.page_source
        html = etree.HTML(page_source)
        result = html.xpath('//html/body/div[3]/div[1]/div/div[2]/div[1]/div[1]/div')
        return result 

    def get_item_info(self, item):
        try:
            title = item.xpath('div/div/div[1]/a/text()')[0]
        except IndexError:
            return 
        if item.xpath('div/div/div[2]/span[2]/text()')[0] in ['(评价人数不足)', '(目前无人评价)']:
            score = 0
            eval_num = 0
        else:
            score = float(item.xpath('div/div/div[2]/span[2]/text()')[0])
            eval_num = int(re.findall('(\d+)', 
                item.xpath('div/div/div[2]/span[3]/text()')[0])[0])
        try:
            meta_abstract = item.xpath('div/div/div[3]/text()')[0]
        except IndexError:
            meta_abstract = ''
        return [title, score, eval_num, meta_abstract]
    
    def is_next_page(self, item):
        if not item.xpath('a'):
            return False
        else:
            if item.xpath('a[last()]/@class')[0] == 'next':
                return True
            else:
                return False

    def rank_res(self, rank_num=1):
        return sorted(self.res_list, key=lambda s:s[1], reverse=True)

    def main(self):
        self.search()
        is_continue = True
        while is_continue:
            items = self.get_items()
            for i, item in enumerate(items):
                try:
                    item_res = self.get_item_info(item)
                    if item_res:
                        self.res_list.append(item_res)
                except IndexError:
                    if i < len(items) - 1:
                        continue
                if i == len(items) - 1:
                    is_next_page = self.is_next_page(item)
                    if is_next_page:
                        xpath = '/html/body/div[3]/div[1]/div/div[2]/div[1]/div[1]/div[last()]/a[last()]'
                        ele = self.driver.find_element_by_xpath(xpath)
                        self.click(ele)
                    else:
                        is_continue = False

        ranked_res = self.rank_res()
        df = pd.DataFrame(ranked_res, columns=['书名', '评分', '评论人数', '作者及出版社情况'])
        path = self.get_path()
        df.to_csv(path, index=False)
        self.driver.quit()
         
if __name__ == '__main__':
    key_word = input("请输入需要查询的书名：")
    crawler = Crawler(key_word)
    crawler.main()
    

