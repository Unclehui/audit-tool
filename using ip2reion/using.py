# -*- coding: utf-8 -*-
"""
Created on Thu Apr 21 15:42:19 2022
IP信息查询：分内网IP、国内一级行政区、国外国家
@author: HJY
"""

from ip2Region import Ip2Region
import pandas as pd

search = Ip2Region('ip2region.db')
def ipsearch (ip):
    re = search.memorySearch(ip)['region'].decode('utf-8').split('|')
    if re[-1] == '内网IP':
        return re[-1]
    elif re[0] == '中国':
        return re[2]
    else:
        return re[0]

df = pd.read_csv('test.csv')
df['region'] = df.apply(lambda x: ipsearch(x.ip), axis = 1)

df.to_csv('result.csv')