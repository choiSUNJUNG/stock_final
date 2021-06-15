# 일 평균 30만주이상 거래되는 nasdaq, newyork, amex(300k_day_coms.xlsx) 중 
# 60일 거래량 평균 대비 60일내 일일 최대 거래량이 500% 이상인 종목(5_volume_follow.xlsx) 중
# macd > 0 인 종목(5_volume_follow_w_macd.xlsx) 중
# 현 주가가 4주(20일선) 전고점 돌파하고 60일선 위에 있고 60일선도 상승중인 종목 추출(step5_volume_trend_follow.xlsx)
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
sheet.append(['time', 'market', 'symbol', 'code', 'company_name', 'price', 'vol_max', 'vol_mean', 'industry', 'macd', 'sto', 'adx'])
wb.save('step5_volume_trend_follow.xlsx')

#회사 데이터 읽기

df_1 = pd.read_excel("step4_volume_follow_w_macd.xlsx") 
df_com = df_1.sort_values(by = 'macd') # macd 기준 올림차순으로 sorting
now = datetime.datetime.now()
i = 1
for i in range(len(df_com)):
    df = pdr.get_data_yahoo(df_com.iloc[i]['symbol'], period = '70d')  # 기간 10일
    df['high_max'] = df['Close'].rolling(window=20).max()
    df['ma60'] = df['Close'].rolling(window=60).mean()
    df['min_14'] = df['Low'].rolling(14).min()
    df['max_14'] = df['High'].rolling(14).max()
    df.min = df['min_14']
    df.max = df['max_14']
    df['sto_K_14'] = (df['Close'] - df.min) / (df.max - df.min) * 100   #stochastics_fast의 %K
    df['sto_D_5'] = df['sto_K_14'].rolling(5).mean()    #stochastics_slow의 %K
    df['sto_DS_3'] = df['sto_D_5'].rolling(3).mean()    #stochastics_slow의 %D
   
    # print(df)
    # 오늘 종가(현재가)가 20 고점(전고점) 보다 높고 60이평선 위에 있으며 60 이평선 또한 상승하는 경우
    if df.iloc[-1]['Close'] > df.iloc[-2]['high_max'] and df.iloc[-1]['Close'] > df.iloc[-1]['ma60'] > df.iloc[-3]['ma60']:  
        sheet.append([now, df_com.iloc[i]['market'], df_com.iloc[i]['symbol'], df_com.iloc[i]['code'], \
            df_com.iloc[i]['company_name'], df_com.iloc[i]['price'], df_com.iloc[i]['vol_max'], df_com.iloc[i]['vol_mean'], \
                df_com.iloc[i]['industry'], df_com.iloc[i]['macd'], df.iloc[-1]['sto_D_5'], 'adx'])
        wb.save('step5_volume_trend_follow.xlsx')
        print('매수발생 : ', df_com.iloc[i]['symbol'])
    
    i += 1  
