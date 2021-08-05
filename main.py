# coding = utf-8

import os
import logging
import configparser
import datetime

import xlwings as xw
import hashlib
import json
import random
from time import sleep, time
import requests
import xlrd
import xlwt
from xlutils.copy import copy

class KuaiDi100:
    def __init__(self):
        self.autourl = 'https://m.kuaidi100.com/apicenter/kdquerytools.do?method=autoComNum&text={}'  # 请求地址
        self.queryurl = 'https://m.kuaidi100.com/query'  # 请求地址
        self.headers = {'Accept': 'application/json, text/javascript, */*; q=0.01',
                'Accept-Encoding': 'gzip, deflate, br',
                'Accept-Language': 'zh-CN,zh;q=0.9',
                'Connection': 'keep-alive',
                'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
                'sec-ch-ua-mobile': '?0',
                'Sec-Fetch-Mode': 'cors',
                'Sec-Fetch-Site': 'same-origin',
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/92.0.4515.107 Safari/537.36',
                'X-Requested-With': 'XMLHttpRequest',
        }
        self.s = requests.Session()
        re = self.s.get('https://m.kuaidi100.com/result.jsp?nu=',headers=self.headers)
        print(re.cookies)

    def auto_number(self, num):
        """
        智能单号识别
        :param num: 快递单号
        :return: requests.Response.text
        """
        
        req_params = {'token': '','platform': 'MWWW'}
        url = self.autourl.format(num)
        re = self.s.post(url, req_params, headers = self.headers)  # 发送请求
        print(re.cookies)
        return re.text

    def tracktmp(self,com,num):
        tmp = random.random()
        param = {
            'postid': num,
            'id': '1',
            'valicode': '',
            'temp': tmp,
            'type': com,
            'phone': '', 
            'token': '',
            'platform': 'MWWW'
        }

        re =  self.s.post(self.queryurl, param, headers = self.headers)  # 发送请求
        print(re.cookies)
        return re.text

 

if os.path.exists("my.log"):
    os.remove("my.log")


LOG_FORMAT = "%(asctime)s - %(levelname)s - %(message)s"
logging.basicConfig(filename='my.log', level=logging.DEBUG, format=LOG_FORMAT)

try:
    file = open('cfg.txt','rb')
    str1 = file.read()
    file.close()
    data = json.loads(str1)
    logging.debug(data)
except OSError as reason:
    logging.debug('open file error:{}'.format(str(reason)))
    exit(0)


# 创建管理对象
# 读ini文件
actsheet = data['tablename'].format(datetime.datetime.now().month,datetime.datetime.now().day)
logging.debug('actsheet:{}'.format(actsheet))
pwd = data['excelname']
logging.debug('excelpwd:{}'.format(pwd))
orderList = data['orderList']
app = xw.App(visible=True,add_book=False)
wb = app.books.open(pwd)
table = wb.sheets[actsheet]
table.activate()
logging.debug('exel表的名字:{}'.format(table.name))
sht = wb.sheets.active
count = wb.sheets.active.range(orderList).expand('table').rows.count
logging.debug('物流单号数量:{}'.format(count))
kd100 = KuaiDi100()
for i in range(count):
    if i != 0:
        rang = sht.range(i+1,1)
        result = kd100.auto_number(rang.value)
        data = json.loads(result)
        logging.debug('物流单号:{} 物流公司:{}'.format(rang.value,data['auto'][0]['name']))
        result = kd100.tracktmp(data['auto'][0]['comCode'], rang.value)
        data = json.loads(result)
        for k in data['data']:
            sht.range(i+1,2).value = k['time']
            sht.range(i+1,3).value = k['context']
            logging.debug('物流订单时间:{} 物流订单状态:{}'.format(k['time'],k['context']))
            break
        sleep(2);
    

wb.save()
wb.close()
app.quit()