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
# sheet.append(['time', 'market', 'symbol', 'code', 'company_name', 'bol_gap(%)', 'rise_margin(%)', 'year_rise(%)', 'price', 'industry', 'trade'])
sheet.append(['time', 'market', 'symbol', 'code', 'company_name', 'year_begin_price', 'year_rise(%)', '6month_price', '6month_rise(%)', \
    '3month_price', '3month_rise(%)','monthly_price', 'monthly_rise(%)','yesterday_price', 'new_price', 'day_rise(%)', 'industry'])
wb.save('f_rise_base_'+filename+'.xlsx')

#회사 데이터 읽기
# df_com = pd.read_excel('step3_3mon_10p_up_2021-07-01') #볼린저밴드 하단에서 상승하는 종목 중 년간 상승률 순서대로 나열하기
df_com = pd.read_excel('step2_300k_day_coms.xlsx')
# print(df_com)
# len(df_com)
# print(len(df_com))

# while True:
#     try:
#         time.sleep(10)
i = 1
for i in range(len(df_com)):
    # now = datetime.datetime.now()
    df = pdr.get_data_yahoo(df_com.iloc[i]['symbol'], period = '12mo')
    try : 
        df.year_begin_price = df.iloc[1]['Close']
        df.half_year_begin_price = df.iloc[120]['Close']
        df.three_mon_price = df.iloc[-61]['Close']
        df.one_mon_price = df.iloc[-21]['Close']
        df.prev_day_price = df.iloc[-2]['Close']
        df.new_price = df.iloc[-1]['Close']
        df.year_rise = (df.new_price / df.year_begin_price - 1) * 100
        df.half_mon_rise = (df.new_price / df.half_year_begin_price - 1) * 100
        df.three_mon_rise = (df.new_price / df.three_mon_price - 1) * 100
        df.mon_rise = (df.new_price / df.one_mon_price - 1) * 100
        df.day_rise = (df.new_price / df.prev_day_price - 1) * 100
        # df['year_rise'] = df.year_rise
        if df.year_rise < 0 and df.three_mon_rise > 0:
            sheet.append([now, df_com.iloc[i]['market'], df_com.iloc[i]['symbol'], df_com.iloc[i]['code'], df_com.iloc[i]['company_name'], \
                df.year_begin_price, df.year_rise, df.half_year_begin_price, df.half_mon_rise, df.three_mon_price, df.three_mon_rise, df.one_mon_price, \
                    df.mon_rise, df.prev_day_price, df.new_price, df.day_rise, df_com.iloc[i]['industry']])
        # sheet.append([now, df_com.iloc[i]['market'], df_com.iloc[i]['symbol'], df_com.iloc[i]['code'], df_com.iloc[i]['company_name'], \
            # df_com.iloc[i]['bol_gap(%)'], df_com.iloc[i]['rise_margin(%)'], df.year_rise, df_com.iloc[i]['price'], df_com.iloc[i]['industry'],df_com.iloc[i]['trade']])
        wb.save('f_rise_base_'+filename+'.xlsx')
                      
            # elif df_b <= 20:
            #     sheet.append([now, df_com.iloc[i]['simbol'], df_com.iloc[i]['company_name'], 'narrow'])
            #     wb.save('watch_data.xlsx')
    except Exception as e:
        print(e)
        print('error', df_com.iloc[i]['symbol'])
    print(df_com.iloc[i]['symbol'] )
    print(i)
    print(len(df_com))
    i += 1   

# df_1 = pd.read_excel('f_rise_base_'+filename+'.xlsx') 
# df_b_f = df_1.sort_values(by = 'year_rise(%)', ascending= False) # gap_close_ratio(%) 기준 올림차순으로 sorting
# df_b_f6 = df_1.sort_values(by = '6month_rise(%)', ascending= False)
# df_b_f3 = df_1.sort_values(by = '3month_rise(%)', ascending= False)
# df_b_f1 = df_1.sort_values(by = 'monthly_rise(%)', ascending= False)
# df_b_f0 = df_1.sort_values(by = 'day_rise(%)', ascending= False)
# df_b_f.to_excel('f_year_rise_base_'+filename+'.xlsx')
# df_b_f6.to_excel('f_6month_rise_base_'+filename+'.xlsx')
# df_b_f3.to_excel('f_3month_rise_base_'+filename+'.xlsx')
# df_b_f1.to_excel('f_month_rise_base_'+filename+'.xlsx')
# df_b_f0.to_excel('f_day_rise_base_'+filename+'.xlsx') 
# df_b_f.to_excel('imsi_trend_bollinger_follow_sorted_'+filename+'.xlsx')
    #         time.sleep(0.1)
    # except Exception as e:
    #     print(e)
    #     time.sleep(0.1)
