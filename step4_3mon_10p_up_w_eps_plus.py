# 순이익 흑자이며 매출액 증가하는 회사(최근 3개월 동안 10%이상 상승한 회사 대상)
# per 20이하, EPS 1.0 이상 찾기
# import pandas_datareader.data as web
from pandas_datareader import data as pdr
import datetime 
import csv
import urllib
from bs4 import  BeautifulSoup as bs
import openpyxl
import time
import yfinance as yf
import pandas as pd

yf.pdr_override()
# import matplotlib.pyplot as plt
wb = openpyxl.Workbook()
now = datetime.datetime.now()
filename = datetime.datetime.now().strftime("%Y-%m-%d")
sheet = wb.active

# cell name 생성
sheet.append(['time', 'market', 'symbol', 'code', 'company_name', '2020Revenue', '2020Earnings', '1Q2021_Revenue', '1Q2021_Earnings', 'industry'])
wb.save('step4_3mon_up_w_eps_plus_'+filename+'.xlsx')

#회사 데이터 읽기
df_com = pd.read_excel("step3_3mon_10p_up_2021-10-26.xlsx")
i = 1
for i in range(len(df_com)):
    # df = pdr.get_data_yahoo(df_com1.iloc[i]['Symbol'], period = '1mo')  # 기간 1month
    # df = pdr.get_data_yahoo(df_com.iloc[i]['symbol'], period = '6mo')  # 기간 10일
    yf_symbol = df_com.iloc[i]['symbol']
    yf_ticker = yf.Ticker(yf_symbol)
    df1 = yf_ticker.earnings
    df2 = yf_ticker.quarterly_earnings
    df3 = pdr.get_data_yahoo(df_com.iloc[i]['symbol'], period = '3d')  # 기간 3일
    # df1 = df_com.iloc[i]['symbol'].earnings
    # df2 = df_com.iloc[i]['symbol'].quarterly_earnings
    try : 
        if df1.iloc[-1]['Earnings'] and df2.iloc[-1]['Earnings'] > 0:
            sheet.append([now, df_com.iloc[i]['market'], df_com.iloc[i]['symbol'], df_com.iloc[i]['code'], df_com.iloc[i]['company_name'], df1.iloc[-1]['Revenue'], \
                df1.iloc[-1]['Earnings'], df2.iloc[-1]['Revenue'], df2.iloc[-1]['Earnings'], df_com.iloc[i]['industry']])
            print(yf_symbol, df3.iloc[-1]['Close'], df1.iloc[-1]['Earnings'],df2.iloc[-1]['Earnings'])
    except Exception as e:
        print(e)
        print('error', df_com.iloc[i]['symbol'])    
    i += 1
wb.save('step4_3mon_up_w_eps_plus_'+filename+'.xlsx')
