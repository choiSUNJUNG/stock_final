# s1_trend_follow.py 실행결과로 얻어진 s1_trend_follow.xlsx'파일을 's2_trend_follow_tot.xlsx'에 추가하고
# 중복된 종목은 최근 데이터만 남기는 과정)
from pandas_datareader import data as pdr
# from openpyxl import load_workbook, Workbook
import yfinance as yf
import pandas as pd
import matplotlib.pyplot as plt
import openpyxl
import datetime
import time

yf.pdr_override()
wb = openpyxl.Workbook()
sheet = wb.active
# cell name 생성
now = datetime.datetime.now()
df1 = pd.read_excel("s2_trend_follow_tot.xlsx")
sheet.append(['time', 'market', 'symbol', 'code', 'company_name', 'price', 'high_max', 'vol_max', 'vol_mean', 'industry', '20or60', 'power'])
i = 1
for i in range(len(df1)) :
    sheet.append([now, df1.iloc[i]['market'], df1.iloc[i]['symbol'], df1.iloc[i]['code'], \
        df1.iloc[i]['company_name'], df1.iloc[i]['price'], df1.iloc[i]['high_max'], df1.iloc[i]['vol_max'], \
            df1.iloc[i]['vol_mean'], df1.iloc[i]['industry'], df1.iloc[i]['20or60'], df1.iloc[i]['power']])
wb.save("s2_trend_follow_tot.xlsx")
df = pd.read_excel("s1_trend_follow.xlsx")
j = 1
for j in range(len(df)) :
    sheet.append([now, df.iloc[j]['market'], df.iloc[j]['symbol'], df.iloc[j]['code'], \
        df.iloc[j]['company_name'], df.iloc[j]['price'], df.iloc[j]['high_max'], df.iloc[j]['vol_max'], \
            df.iloc[j]['vol_mean'], df.iloc[j]['industry'], df.iloc[j]['20or60'], df.iloc[j]['power']])
    # wb.save('s1_trend_follow_tot.xlsx')
wb.save("s2_trend_follow_tot.xlsx")

#중복 데이터 최신데이터만 남기고 삭제
df_com = pd.read_excel("s2_trend_follow_tot.xlsx")
df_com1 = df_com.drop_duplicates(['symbol'], keep = 'last')
df_com1.to_excel('s2_trend_follow_tot.xlsx')
