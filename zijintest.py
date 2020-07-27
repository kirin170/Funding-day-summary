#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Time : 2019/12/31 10:35
# @Author : Kirin
# @Site : 
# @File : zijintest.py
# @Software: PyCharm

import pandas as pd
from openpyxl import load_workbook
from enum import Enum

class Company(Enum):
    JIA = '甲'
    YI = '乙'
    BING = '丙'
    DING = '丁'
    WU = '戊'

dir_end = '日报汇总表\\最终资金汇总.xlsx'

jia_zuoriyue = []
jia_jinrichengdui = []
jia_jinrihejiyue = []

df = pd.read_excel('日报汇总表\资金汇总12.27.xlsx', sheet_name='资金汇总11')
df.columns = ['公司名称', '昨日余额', '银行存款', '银行承兑', '小计', '备注']
df.set_index([df.columns[0]], inplace=True)

#pandas写入excel多个sheet不覆盖
book = load_workbook(dir_end)
#xls用xlwt。xlsx用openpyxl
writer = pd.ExcelWriter(dir_end, engine='openpyxl')
writer.book = book
writer.sheets = dict((ws.title, ws) for ws in book.worksheets)

for i in range(1, 32):
    jia_day = pd.read_excel('日报汇总表\\丙.xlsx', sheet_name=str(i))
    jia_day.fillna(0, inplace=True)
    col_name = jia_day.iloc[1, 0:].tolist()
    jia_day.columns = col_name
    jia_day.set_index([col_name[0]], inplace=True)

    jia_zuoriyue.append(round(jia_day.loc['三、日合计', '昨日余额'], 2))
    jia_jinrichengdui.append(round(jia_day.loc['二、承兑', '今日余额'], 2))
    jia_jinrihejiyue.append(round(jia_day.loc['三、日合计', '今日余额'], 2))

    df.loc['丙', '昨日余额'] = jia_zuoriyue[i-1]
    df.loc['丙', '银行承兑'] = jia_jinrichengdui[i-1]
    df.loc['丙', '小计'] = jia_jinrihejiyue[i-1]
    df.loc['丙', '银行存款'] = df.loc['丙', '小计'] - df.loc['丙', '银行承兑']

    df.to_excel(writer, sheet_name=str(i), header=False)
writer.save()
