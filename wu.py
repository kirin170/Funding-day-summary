#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Time : 2019/12/31 13:53
# @Author : Kirin
# @Site : 
# @File : wu.py
# @Software: PyCharm

import pandas as pd

wu_zuoriyue = []
wu_jinrichengdui = []
wu_jinrihejiyue = []
for i in range(1,32):

    wu_day = pd.read_excel('日报汇总表\戊.xls', sheet_name = str(i))
    wu_day.fillna(0,inplace=True)
    #print(bing_day1)
    col_name = wu_day.iloc[1,0:].tolist()
    #row_name = bing_day1.iloc[0:,0].tolist()

    wu_day.columns = col_name
    wu_day.set_index([col_name[0]], inplace=True)

    wu_zuoriyue.append(round(wu_day.loc['三、日合计','昨日余额'],2))
    wu_jinrichengdui.append(round(wu_day.loc['二、承兑','今日余额'],2))
    wu_jinrihejiyue.append(round(wu_day.loc['三、日合计','今日余额'],2))
print('昨日余额日合计：')
print(wu_zuoriyue)
print('今日承兑余额：')
print(wu_jinrichengdui)
print('今日合计余额：')
print(wu_jinrihejiyue)