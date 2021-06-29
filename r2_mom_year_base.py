# 종목 파일에서 1년 상대수익률 계산해서 순서대로 나열하기
from pandas_datareader import data as pdr
import yfinance as yf
import pandas as pd
import matplotlib.pyplot as plt
import openpyxl
import datetime
import time

yf.pdr_override()
wb = openpyxl.Workbook()
now = datetime.datetime.now()
filename = datetime.datetime.now().strftime("%Y-%m-%d")

# wb.save('watch_data.xlsx')
sheet = wb.active
# cell name : date, simbol, company name, upper or lower or narrow band 
sheet.append(['time', 'market', 'symbol', 'code', 'company_name', 'bol_gap(%)', 'rise_margin(%)', 'year_rise(%)', 'price', 'industry', 'trade'])
wb.save('r2_mom_year_base_'+filename+'.xlsx')

#회사 데이터 읽기
df_com = pd.read_excel('rel_mom_bollinger_follow_'+filename+'.xlsx')
# print(df_com)
# len(df_com)
# print(len(df_com))

# while True:
#     try:
#         time.sleep(10)
i = 1
for i in range(len(df_com)):
    # now = datetime.datetime.now()
    df = pdr.get_data_yahoo(df_com.iloc[i]['symbol'], period = '3mo')
    df.old_price = df.iloc[1]['Close']
    df.new_price = df.iloc[-1]['Close']
    df.year_rise = (df.new_price / df.old_price - 1) * 100
    # df['year_rise'] = df.year_rise
    sheet.append([now, df_com.iloc[i]['market'], df_com.iloc[i]['symbol'], df_com.iloc[i]['code'], df_com.iloc[i]['company_name'], \
        df_com.iloc[i]['bol_gap(%)'], df_com.iloc[i]['rise_margin(%)'], df.year_rise, df_com.iloc[i]['price'], df_com.iloc[i]['industry'],df_com.iloc[i]['trade']])
    wb.save('r2_mom_year_base_'+filename+'.xlsx')
                      
            # elif df_b <= 20:
            #     sheet.append([now, df_com.iloc[i]['simbol'], df_com.iloc[i]['company_name'], 'narrow'])
            #     wb.save('watch_data.xlsx')
    i += 1   

df_1 = pd.read_excel('r2_mom_year_base_'+filename+'.xlsx') 
df_b_f = df_1.sort_values(by = 'year_rise(%)', ascending= False) # gap_close_ratio(%) 기준 올림차순으로 sorting
df_b_f.to_excel('r2_mom_year_base_'+filename+'.xlsx')
# df_b_f.to_excel('imsi_trend_bollinger_follow_sorted_'+filename+'.xlsx')
    #         time.sleep(0.1)
    # except Exception as e:
    #     print(e)
    #     time.sleep(0.1)
