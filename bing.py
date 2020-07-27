#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Time : 2019/12/31 10:41
# @Author : Kirin
# @Site : 
# @File : bing.py
# @Software: PyCharm
import pandas as pd

bing_zuoriyue = []
bing_jinrichengdui = []
bing_jinrihejiyue = []
for i in range(1,32):

    bing_day = pd.read_excel('日报汇总表\丙.xlsx', sheet_name = str(i))
    bing_day.fillna(0,inplace=True)
    #print(bing_day1)
    col_name = bing_day.iloc[1,0:].tolist()
    #row_name = bing_day1.iloc[0:,0].tolist()

    bing_day.columns = col_name
    bing_day.set_index([col_name[0]], inplace=True)
    # data1 = round(bing_day1.loc['三、日合计','昨日余额'],2)
    # data2 = round(bing_day1.loc['二、承兑','今日余额'],2)
    # data3 = round(bing_day1.loc['三、日合计','今日余额'],2)
    bing_zuoriyue.append(round(bing_day.loc['三、日合计','昨日余额'],2))
    bing_jinrichengdui.append(round(bing_day.loc['二、承兑','今日余额'],2))
    bing_jinrihejiyue.append(round(bing_day.loc['三、日合计','今日余额'],2))
print('昨日余额日合计：')
print(bing_zuoriyue)
print('今日承兑余额：')
print(bing_jinrichengdui)
print('今日合计余额：')
print(bing_jinrihejiyue)

