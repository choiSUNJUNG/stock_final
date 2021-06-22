# 일 평균 30만주이상 거래되는 nasdaq, newyork, amex(300k_day_coms.xlsx) 중 
# 현 주가가 12주(60일) 전고점 돌파하고 상승하는 120일선 위에 있는 종목
# (120일 거래량 평균 대비 120일내 일일 최대 거래량이 500% 이상인 경우 '●power'로 표시
# 
# 추출종목 's1_trend_follow.xlsx'은 's1_trend_follow_tot.xlsx'로 추가함(수동)
# 
# 
from pandas_datareader import data as pdr
import yfinance as yf
import pandas as pd
import openpyxl
import datetime
import time

yf.pdr_override()
wb = openpyxl.Workbook()
sheet = wb.active
# cell name 생성
sheet.append(['time', 'market', 'symbol', 'code', 'company_name', 'price', 'high_max', 'vol_max', 'vol_mean', 'industry', '20or60', 'power'])
wb.save('s1_trend_follow.xlsx')

#회사 데이터 읽기

df_com = pd.read_excel("step2_300k_day_coms.xlsx") 
now = datetime.datetime.now()
i = 1
for i in range(len(df_com)):
    df = pdr.get_data_yahoo(df_com.iloc[i]['symbol'], period = '150d')  # 기간 6개월
    df['high_20_max'] = df['Close'].rolling(window=20).max()
    df['high_60_max'] = df['Close'].rolling(window=60).max()
    df['ma120'] = df['Close'].rolling(window=120).mean()
    df['vol_mean'] = df['Volume'].rolling(window=120).mean()
    df['vol_max'] = df['Volume'].rolling(window=120).max()
    
    # print(df)
    # 오늘 종가(현재가)가 20 또는 60 고점(전고점) 보다 높고 120이평선 위에 있으며 120 이평선 또한 상승하는 경우
    if df.iloc[-1]['Close'] >= df.iloc[-2]['high_60_max'] and df.iloc[-1]['Close'] > df.iloc[-1]['ma120'] > df.iloc[-3]['ma120']:    #60일 최고치
        if df.iloc[-1]['vol_max'] > df.iloc[-1]['vol_mean'] * 5 :   # 120일 거래량 평균 대비 120일내 일일 최대 거래량이 500% 이상인 종목
            sheet.append([now, df_com.iloc[i]['market'], df_com.iloc[i]['symbol'], df_com.iloc[i]['code'], \
                df_com.iloc[i]['company_name'], df.iloc[-1]['Close'], df.iloc[-2]['high_60_max'], df.iloc[-1]['vol_max'], \
                    df.iloc[-1]['vol_mean'], df_com.iloc[i]['industry'], '20', '●power'])
            wb.save('s1_trend_follow.xlsx')
        else : # 거래량 무관
            sheet.append([now, df_com.iloc[i]['market'], df_com.iloc[i]['symbol'], df_com.iloc[i]['code'], \
                df_com.iloc[i]['company_name'], df.iloc[-1]['Close'], df.iloc[-2]['high_60_max'], df.iloc[-1]['vol_max'], \
                    df.iloc[-1]['vol_mean'], df_com.iloc[i]['industry'], '60', ''])
            wb.save('s1_trend_follow.xlsx')
        print('60일 매수발생 : ', df_com.iloc[i]['symbol'], df.iloc[-1]['Close'], df.iloc[-2]['high_60_max'])
    elif df.iloc[-1]['Close'] >= df.iloc[-2]['high_20_max'] and df.iloc[-1]['Close'] > df.iloc[-1]['ma120'] > df.iloc[-3]['ma120']:  #60일 아닌 경우 20일 최고치
        if df.iloc[-1]['vol_max'] > df.iloc[-1]['vol_mean'] * 5 :   # 120일 거래량 평균 대비 120일내 일일 최대 거래량이 500% 이상인 종목
            sheet.append([now, df_com.iloc[i]['market'], df_com.iloc[i]['symbol'], df_com.iloc[i]['code'], \
                df_com.iloc[i]['company_name'], df.iloc[-1]['Close'], df.iloc[-2]['high_20_max'], df.iloc[-1]['vol_max'], \
                    df.iloc[-1]['vol_mean'], df_com.iloc[i]['industry'], '60', '●power'])
            wb.save('s1_trend_follow.xlsx')
        else : # 거래량 무관
            sheet.append([now, df_com.iloc[i]['market'], df_com.iloc[i]['symbol'], df_com.iloc[i]['code'], \
                df_com.iloc[i]['company_name'], df.iloc[-1]['Close'], df.iloc[-2]['high_20_max'], df.iloc[-1]['vol_max'], \
                    df.iloc[-1]['vol_mean'], df_com.iloc[i]['industry'], '20', ''])
            wb.save('s1_trend_follow.xlsx')
        print('20일 매수발생 : ', df_com.iloc[i]['symbol'], df.iloc[-1]['Close'], df.iloc[-2]['high_20_max'])
    i += 1  