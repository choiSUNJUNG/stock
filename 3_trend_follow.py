# 리처드 데니스의 추세추종전략 
# 4주(20일선) 전고점 돌파 매수, 2주(10일선) 전저점 돌파 매도
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
sheet.append(['time', 'market', 'symbol', 'code', 'company_name', 'price', 'pre_high_low', 'volume', 'industry', 'sign'])
wb.save('trend_follow_sign.xlsx')

#회사 데이터 읽기
df_com = pd.read_excel("300k_day_coms.xlsx")
now = datetime.datetime.now()
i = 1
for i in range(len(df_com)):
    # df = pdr.get_data_yahoo(df_com1.iloc[i]['Symbol'], period = '1mo')  # 기간 1month
    df = pdr.get_data_yahoo(df_com.iloc[i]['symbol'], period = '30d')  # 기간 10일
    df['high_max'] = df['High'].rolling(window=20).max()
    df['low_min'] = df['Low'].rolling(window=10).min()
    # print(df)
  
    if df.iloc[-1]['Close'] > df.iloc[-2]['high_max'] and df.iloc[-1]['Close'] < 100.0 :  # 오늘 종가(현재가)가 20거래일 고점(전고점)을 넘어서는 조건
        sheet.append([now, df_com.iloc[i]['market'], df_com.iloc[i]['symbol'], df_com.iloc[i]['code'], \
            df_com.iloc[i]['company_name'], df.iloc[-1]['Close'], df.iloc[-2]['high_max'], df.iloc[-1]['Volume'], \
                df_com.iloc[i]['industry'], 'buy'])
        wb.save('trend_follow_sign.xlsx')
        print('매수발생', df_com.iloc[i]['symbol'])
    elif df.iloc[-1]['Close'] < df.iloc[-2]['low_min'] and df.iloc[-1]['Close'] < 100.0  :  # 오늘 종가(현재가)가 10거래일 저점(전저점)보다 내려오는 조건
        sheet.append([now, df_com.iloc[i]['market'], df_com.iloc[i]['symbol'], df_com.iloc[i]['code'], \
            df_com.iloc[i]['company_name'], df.iloc[-1]['Close'], df.iloc[-2]['low_min'], df.iloc[-1]['Volume'], \
                df_com.iloc[i]['industry'], 'sell'])
        wb.save('trend_follow_sign.xlsx')
        print('매도발생', df_com.iloc[i]['symbol'])
    
    i += 1   