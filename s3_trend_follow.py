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
sheet.append(['time', 'market', 'symbol', 'code', 'company_name', 'price', 'high_max', 'vol_max', 'vol_mean', 'industry', 'power'])
wb.save('s3_trend_follow.xlsx')

#회사 데이터 읽기

df_com = pd.read_excel("step2_300k_day_coms.xlsx") 
now = datetime.datetime.now()
i = 1
for i in range(len(df_com)):
    df = pdr.get_data_yahoo(df_com.iloc[i]['symbol'], period = '150d')  # 기간 6개월
    df['high_max'] = df['Close'].rolling(window=60).max()
    df['ma120'] = df['Close'].rolling(window=120).mean()
    df['vol_mean'] = df['Volume'].rolling(window=120).mean()
    df['vol_max'] = df['Volume'].rolling(window=120).max()
    
    # print(df)
    # 오늘 종가(현재가)가 120 고점(전고점) 보다 높고 120이평선 위에 있으며 60 이평선 또한 상승하는 경우
    if df.iloc[-1]['Close'] >= df.iloc[-2]['high_max'] and df.iloc[-1]['Close'] > df.iloc[-1]['ma120'] > df.iloc[-3]['ma120']:  
        if df.iloc[-1]['vol_max'] > df.iloc[-1]['vol_mean'] * 5 :
            sheet.append([now, df_com.iloc[i]['market'], df_com.iloc[i]['symbol'], df_com.iloc[i]['code'], \
                df_com.iloc[i]['company_name'], df.iloc[-1]['Close'], df.iloc[-2]['high_max'], df.iloc[-1]['vol_max'], \
                    df.iloc[-1]['vol_mean'], df_com.iloc[i]['industry'], '●power'])
            wb.save('s3_trend_follow.xlsx')
        else :
            sheet.append([now, df_com.iloc[i]['market'], df_com.iloc[i]['symbol'], df_com.iloc[i]['code'], \
                df_com.iloc[i]['company_name'], df.iloc[-1]['Close'], df.iloc[-2]['high_max'], df.iloc[-1]['vol_max'], \
                    df.iloc[-1]['vol_mean'], df_com.iloc[i]['industry'], ''])
            wb.save('s3_trend_follow.xlsx')
        print('매수발생 : ', df_com.iloc[i]['symbol'], df.iloc[-1]['Close'], df.iloc[-2]['high_max'])
    
    i += 1  

   