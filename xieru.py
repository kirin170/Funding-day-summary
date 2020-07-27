#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Time : 2020/1/3 15:49
# @Author : Kirin
# @Site : 
# @File : xieru.py
# @Software: PyCharm

import pandas as pd
from openpyxl import load_workbook

jia_zuoriyue = []
jia_jinrichengdui = []
jia_jinrihejiyue = []

df = pd.read_excel('日报汇总表\资金汇总12.27.xlsx', sheet_name='资金汇总11')
df.columns = ['公司名称', '昨日余额', '银行存款', '银行承兑', '小计', '备注']
df.set_index([df.columns[0]], inplace=True)

#pandas写入excel多个sheet不覆盖
book = load_workbook('日报汇总表\\最终资金汇总.xlsx')
#xls用xlwt。xlsx用openpyxl
writer = pd.ExcelWriter('日报汇总表\\最终资金汇总.xlsx', engine='openpyxl')
writer.book = book
writer.sheets = dict((ws.title, ws) for ws in book.worksheets)

for i in range(1, 32):
    jia_day = pd.read_excel('日报汇总表\\甲.xlsx', sheet_name=str(i))
    jia_day.fillna(0, inplace=True)
    #print(bing_day1)
    col_name = jia_day.iloc[1, 0:].tolist()
    jia_day.columns = col_name
    jia_day.set_index([col_name[0]], inplace=True)

    jia_zuoriyue.append(round(jia_day.loc['三、日合计', '昨日余额'], 2))
    jia_jinrichengdui.append(round(jia_day.loc['二、承兑', '今日余额'], 2))
    jia_jinrihejiyue.append(round(jia_day.loc['三、日合计', '今日余额'], 2))

    df.loc['甲', '昨日余额'] = jia_zuoriyue[i-1]
    df.loc['甲', '银行承兑'] = jia_jinrichengdui[i-1]
    df.loc['甲', '小计'] = jia_jinrihejiyue[i-1]
    df.loc['甲', '银行存款'] = df.loc['甲', '小计'] - df.loc['甲', '银行承兑']

    df.to_excel(writer, sheet_name=str(i), header=False)
writer.save()

# print('昨日余额日合计：')
# print(jia_zuoriyue)
# print('今日承兑余额：')
# print(jia_jinrichengdui)
# print('今日合计余额：')
# print(jia_jinrihejiyue)


