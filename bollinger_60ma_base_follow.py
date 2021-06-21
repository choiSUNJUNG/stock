# 일 평균 30만주이상 거래되는 nasdaq, newyork, amex 중(step2_300k_day_coms.xlsx)
# 볼린저밴드 하단 근처 종목 중 당일 종가가 전일 시작가 또는 종가 중 큰 값보다 높은 종목 중
# 60일 이평선이 상승중인 종목 추출(bollinger_follow.xlsx)
# 일일 1회 가동 
from pandas_datareader import data as pdr
import yfinance as yf
import pandas as pd
import matplotlib.pyplot as plt
import openpyxl
import datetime
import time

yf.pdr_override()
wb = openpyxl.Workbook()
# wb.save('watch_data.xlsx')
sheet = wb.active
# cell name : date, simbol, company name, upper or lower or narrow band 
sheet.append(['time', 'market', 'symbol', 'code', 'company_name', 'bol_high', 'bol_lower', 'bol_gap(%)', 'price', 'industry', 'trade'])
wb.save('bollinger_follow.xlsx')

#회사 데이터 읽기
df_com = pd.read_excel("step2_300k_day_coms.xlsx")
# print(df_com)
# len(df_com)
# print(len(df_com))
now = datetime.datetime.now()
# while True:
#     try:
#         time.sleep(10)
i = 1
for i in range(len(df_com)):
    # now = datetime.datetime.now()
    df = pdr.get_data_yahoo(df_com.iloc[i]['symbol'], period = '70d')
    df['MA20'] = df['Close'].rolling(window=20).mean()
    df['MA60'] = df['Close'].rolling(window=60).mean()
    df['stddev'] = df['Close'].rolling(window=20).std()
    df['upper'] = df['MA20'] + (df['stddev']*2)
    df['lower'] = df['MA20'] - (df['stddev']*2)
    # df['vol_avr'] = df['Volume'].rolling(window=5).mean()
    df['gap'] = df['upper'] - df['lower']
    df['bandwidth'] = (df['upper'] - df['lower']) / df['MA20'] * 100
    # df = df[19:]
    cur_price = df.iloc[-1]['Close']
    pre_open_price = df.iloc[-2]['Open']
    a2 = max([df.iloc[-2]['Open'], df.iloc[-2]['Close']])
    a3 = max([df.iloc[-3]['Open'], df.iloc[-3]['Close']])
    df_u = df.iloc[-1]['upper']
    df_l1 = df.iloc[-1]['lower']
    df_l2 = df.iloc[-2]['lower']
    df_l3 = df.iloc[-3]['lower']
    df_g1 = df.iloc[-1]['gap']
    df_g2 = df.iloc[-2]['gap']
    df_g3 = df.iloc[-3]['gap']
    # df_v = df.iloc[-2]['vol_avr']
    # print(df_com.iloc[i]['simbol'])
    # print('볼린저 : ',df.iloc[-1]['MA20'], df.iloc[-1]['upper'], df.iloc[-1]['lower'])
    # print('볼린저 밴드폭 : ',df.iloc[-1]['bandwidth'], '%')


    if ((df_l3 - df_g3* 0.3) < a3 < (df_l3 + df_g3* 0.3)) or ((df_l2 - df_g2* 0.2) < a2 < (df_l2 + df_g2* 0.2)) :
        if (df.iloc[-2]['Close'] > a3 or df.iloc[-1]['Close'] > a2) and df.iloc[-1]['MA60'] > df.iloc[-3]['MA60']:
            sheet.append([now, df_com.iloc[i]['market'], df_com.iloc[i]['symbol'], df_com.iloc[i]['code'], df_com.iloc[i]['company_name'], \
                df.iloc[-1]['upper'], df.iloc[-1]['lower'], df.iloc[-1]['gap'], df.iloc[-1]['Close'], df_com.iloc[i]['industry'],'buy'])
            wb.save('bollinger_follow.xlsx')
            print('buy', df_com.iloc[i]['symbol'])
           
            # elif df_b <= 20:
            #     sheet.append([now, df_com.iloc[i]['simbol'], df_com.iloc[i]['company_name'], 'narrow'])
            #     wb.save('watch_data.xlsx')
    i += 1   
    #         time.sleep(0.1)
    # except Exception as e:
    #     print(e)
    #     time.sleep(0.1)
