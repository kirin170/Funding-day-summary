#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Time : 2019/12/31 11:43
# @Author : Kirin
# @Site : 
# @File : jia.py
# @Software: PyCharm

import pandas as pd

jia_zuoriyue = []
jia_jinrichengdui = []
jia_jinrihejiyue = []
for i in range(1,32):

    jia_day = pd.read_excel('日报汇总表\甲.xlsx', sheet_name = str(i))
    jia_day.fillna(0,inplace=True)
    #print(bing_day1)
    col_name = jia_day.iloc[1,0:].tolist()
    #row_name = bing_day1.iloc[0:,0].tolist()

    jia_day.columns = col_name
    jia_day.set_index([col_name[0]], inplace=True)
    # data1 = round(bing_day1.loc['三、日合计','昨日余额'],2)
    # data2 = round(bing_day1.loc['二、承兑','今日余额'],2)
    # data3 = round(bing_day1.loc['三、日合计','今日余额'],2)
    jia_zuoriyue.append(round(jia_day.loc['三、日合计','昨日余额'],2))
    jia_jinrichengdui.append(round(jia_day.loc['二、承兑','今日余额'],2))
    jia_jinrihejiyue.append(round(jia_day.loc['三、日合计','今日余额'],2))
print('昨日余额日合计：')
print(jia_zuoriyue)
print('今日承兑余额：')
print(jia_jinrichengdui)
print('今日合计余额：')
print(jia_jinrihejiyue)