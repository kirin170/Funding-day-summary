#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Time : 2019/12/31 11:50
# @Author : Kirin
# @Site : 
# @File : ding.py
# @Software: PyCharm

import pandas as pd

ding_zuoriyue = []
ding_jinrichengdui = []
ding_jinrihejiyue = []
for i in range(1,32):

    ding_day = pd.read_excel('日报汇总表\丁.xls', sheet_name = str(i))
    ding_day.fillna(0,inplace=True)
    #print(bing_day1)
    col_name = ding_day.iloc[1,0:].tolist()
    #row_name = bing_day1.iloc[0:,0].tolist()

    ding_day.columns = col_name
    ding_day.set_index([col_name[0]], inplace=True)

    ding_zuoriyue.append(round(ding_day.loc['三、日合计','昨日余额'],2))
    ding_jinrichengdui.append(round(ding_day.loc['二、承兑','今日余额'],2))
    ding_jinrihejiyue.append(round(ding_day.loc['三、日合计','今日余额'],2))
print('昨日余额日合计：')
print(ding_zuoriyue)
print('今日承兑余额：')
print(ding_jinrichengdui)
print('今日合计余额：')
print(ding_jinrihejiyue)