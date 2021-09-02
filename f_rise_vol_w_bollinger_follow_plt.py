# 일 평균 30만주이상 거래되는 nasdaq, newyork, amex 중(300k_day_coms.xlsx)
# 직전 10거래일 기준 거래량이 평균 대비 일일 200% 이상 발생한 종목 또는 
# 직전 5거래일 기준 종가가 평균 종가 대비 일일 10% 이상 상승 발생한 종목(f_rise_vol.xlsx) 중에서
# 볼린저밴드가 네로우인 조건 충족하는 종목 추출(f_rise_vol_w_bollinger.xlsx)
# 최근 3년간 매출액 증가하고 순이익 흑자이면서 전년대비 증가한 회사(최근 3개월 동안 10%이상 상승한 회사 대상) 인지 별도 체크
# plot으로 확인
# 일1회 가동
from pandas_datareader import data as pdr
import yfinance as yf
import pandas as pd
import matplotlib.pyplot as plt
import mplfinance as mpf
import openpyxl
import datetime
import time

yf.pdr_override()

#회사 데이터 읽기
df_com = pd.read_excel("f_rise_vol_2021-09-02.xlsx")
i = 1
for i in range(len(df_com)):
    # now = datetime.datetime.now()
    df = pdr.get_data_yahoo(df_com.iloc[i]['symbol'], period = '70d')
    try :
        df['MA20'] = df['Close'].rolling(window=20).mean()
        df['MA60'] = df['Close'].rolling(window=60).mean()
        df['stddev'] = df['Close'].rolling(window=20).std()
        df['upper'] = df['MA20'] + (df['stddev']*2)
        df['lower'] = df['MA20'] - (df['stddev']*2)
        # monthly_rise = df_com.iloc[-1]['6month_rise']
        # df['vol_avr'] = df['Volume'].rolling(window=5).mean()
        # df['gap'] = df['upper'] - df['lower']
        # df['rise_margin'] = (df['upper'] - df['Close']) / df['Close'] * 100
       
        plt.figure(figsize=(9, 7))
        plt.subplot(2, 1, 1)
        plt.plot(df['upper'], color='green', label='Bollinger upper')
        plt.plot(df['lower'], color='brown', label='Bollinger lower')
        plt.plot(df['MA20'], color='red', label='MA20')
        plt.plot(df['MA60'], color='black', label='MA60')
        plt.plot(df['Close'], color='blue', label='Price')
        
        # plt.text(df_com.iloc[i]['company_name'])     
        plt.title(df_com.iloc[i]['symbol']+'remark')
        plt.xlabel('time')
        plt.xticks(rotation = 45)
        plt.ylabel('stock price')
        plt.legend()
            
        plt.subplot(2, 1, 2)
        plt.title(df_com.iloc[i]['remark'])
        plt.plot(df['Volume'], color='blue', label='Volume')
        plt.ylabel('Volume')
        plt.xlabel('time')
        plt.xticks(rotation = 45)
        plt.legend()
        plt.show() 
        # df = df[['Open', 'High', 'Close', 'Volume']] 
                  
    except Exception as e:
        print(e)
        print('error', df_com.iloc[i]['symbol'])
    print(df_com.iloc[i]['symbol'] )
    print(i)
    print(len(df_com))
    i += 1   