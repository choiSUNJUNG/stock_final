# 일 평균 30만주이상 거래되는 nasdaq, newyork, amex 중 
# 60일 거래량 평균 대비 60일내 일일 최대 거래량이 500% 이상인 종목(5_volume_follow.xlsx) 중
# macd > 0 인 종목 추출
# 일 1회 가동!!!(macd 0선 위 올라오는 신규종목 도출)(step4_volume_follow_w_macd.xlsx)

from pandas_datareader import data as pdr
import yfinance as yf
import pandas as pd
import openpyxl
import datetime
import time

def macd(stick, period):
    df1 = pdr.get_data_yahoo(stick, start=period) 

    # 12일 EMA = EMA12
    if len(df1.Close) < 26:
        print("Stock info is short")

    # y = df.Close.values[0]
    # m_list = [y]
    y1 = df1.iloc[0]['Close']
    y2 = df1.iloc[0]['Close']
    m_list1 = [y1] 
    m_list2 = [y2]
    # 12일 지수이동평균 계산
    for i in range(len(df1)):
        if i < 12:
            a12 = 2 / (i+1 + 1)
        else:
            a12 = 2 / (12 + 1)
        y1 = y1*(1-a12) + df1.Close.values[i]*a12
        m_list1.append(y1)

    # 26일 지수이동평균 계산    
    for k in range(len(df1)):
        if k < 26:
            a26 = 2 / (k+1 + 1)
        else:
            a26 = 2 / (26 + 1)
        y2 = y2*(1-a26) + df1.Close.values[k]*a26
        m_list2.append(y2)
    
   
    # macd 계산
    # print(m_list1)
    # print(m_list2)
    macd = m_list1[-1] - m_list2[-1]
    # print(macd)
    return macd

yf.pdr_override()
wb = openpyxl.Workbook()
sheet = wb.active
# cell name 생성
sheet.append(['time', 'market', 'symbol', 'code', 'company_name', 'price', 'vol_max', 'vol_mean', 'industry', 'macd'])
wb.save('step4_volume_follow_w_macd.xlsx')

#회사 데이터 읽기
df_com = pd.read_excel("step3_volume_follow.xlsx")
now = datetime.datetime.now()
start_day = '2021-01-01'

i = 1
for i in range(len(df_com)):
    # df = pdr.get_data_yahoo(df_com1.iloc[i]['Symbol'], period = '1mo')  # 기간 1month
    df = pdr.get_data_yahoo(df_com.iloc[i]['symbol'], period = '70d')  # 기간 10일
    df['vol_mean'] = df['Volume'].rolling(window=60).mean()
    df['vol_max'] = df['Volume'].rolling(window=60).max()
    df['high_max_60'] = df['Close'].rolling(window=60).max()
    
    # print(df)
  
    macd(df_com.iloc[i]['symbol'], start_day)
    if macd(df_com.iloc[i]['symbol'], start_day) > 0 :
        print(macd(df_com.iloc[i]['symbol'], start_day))
        sheet.append([now, df_com.iloc[i]['market'], df_com.iloc[i]['symbol'], df_com.iloc[i]['code'], \
            df_com.iloc[i]['company_name'], df.iloc[-1]['Close'], df.iloc[-1]['vol_max'], df.iloc[-1]['vol_mean'],\
                df_com.iloc[i]['industry'], macd(df_com.iloc[i]['symbol'], start_day)])
        wb.save('step4_volume_follow_w_macd.xlsx')
        print('macd 0선 위', df_com.iloc[i]['symbol'])
    
    i += 1   
