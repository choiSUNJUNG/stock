# 리처드 데니스의 추세추종전략 
# 보유종목 2주(10일선) 전저점 돌파 Check
from pandas_datareader import data as pdr
import yfinance as yf
import pandas as pd
import openpyxl
import datetime
import time

#보유종목 데이터 읽기
yf.pdr_override()
df_com = pd.read_excel("cur_holding_stock.xlsx")
now = datetime.datetime.now()
i = 1
for i in range(len(df_com)):
    # df = pdr.get_data_yahoo(df_com1.iloc[i]['Symbol'], period = '1mo')  # 기간 1month
    df = pdr.get_data_yahoo(df_com.iloc[i]['symbol'], period = '15d')  # 기간 10일
    df['low_min'] = df['Low'].rolling(window=10).min()
    df['pur_price'] = df_com.iloc[i]['pur_price']
    df['profit'] = (df.iloc[-1]['Close'] - df['pur_price']) / df['pur_price'] * 100
        # print(df)
  
    if df.iloc[-1]['Close'] < df.iloc[-2]['low_min'] :  # 오늘 종가(현재가)가 10거래일 저점(전저점)보다 내려오는 조건
        print('매도신호 발생', df_com.iloc[i]['symbol'], '구입가 : ', df.iloc[-1]['pur_price'], '현재가 : ', df.iloc[-1]['Close'], \
            '10일 최저가 : ', df.iloc[-2]['low_min'], '수익률 : ', df.iloc[-1]['profit'], '%')
    else :
        print('유지', df_com.iloc[i]['symbol'], '구입가 : ', df.iloc[-1]['pur_price'], '현재가 : ', df.iloc[-1]['Close'], \
            '10일 최저가 : ', df.iloc[-2]['low_min'], '수익률 : ', df.iloc[-1]['profit'], '%')
    i += 1   