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
wb.save('trend_follow_60.xlsx')

#회사 데이터 읽기
# df_com = pd.read_excel("300k_day_coms.xlsx")
df_com = pd.read_excel("300k_day_coms.xlsx")
now = datetime.datetime.now()
i = 1
for i in range(len(df_com)):
    # df = pdr.get_data_yahoo(df_com1.iloc[i]['Symbol'], period = '1mo')  # 기간 1month
    df = pdr.get_data_yahoo(df_com.iloc[i]['symbol'], period = '70d')  # 기간 10일
    df['high_max'] = df['Close'].rolling(window=20).max()
    df['high_max_60'] = df['Close'].rolling(window=60).max()
    df['low_min'] = df['Close'].rolling(window=10).min()
    # print(df)

    if df.iloc[-2]['Close'] < df.iloc[-3]['high_max_60'] and df.iloc[-1]['Close'] > df.iloc[-2]['high_max_60']: 
        # 60일 전고점 돌파 첫날
        sheet.append([now, df_com.iloc[i]['market'], df_com.iloc[i]['symbol'], df_com.iloc[i]['code'], \
            df_com.iloc[i]['company_name'], df.iloc[-1]['Close'], df.iloc[-2]['high_max'], df.iloc[-1]['Volume'], \
                df_com.iloc[i]['industry'], '60overday'])
        wb.save('trend_follow_60.xlsx')
        print('60일 전고점 돌파 첫날', df_com.iloc[i]['symbol'])
    elif df.iloc[-2]['Close'] > df.iloc[-3]['high_max_60'] and df.iloc[-1]['Close'] > df.iloc[-2]['high_max_60'] : # 60일 전고점 위 상승 지속
        sheet.append([now, df_com.iloc[i]['market'], df_com.iloc[i]['symbol'], df_com.iloc[i]['code'], \
            df_com.iloc[i]['company_name'], df.iloc[-1]['Close'], df.iloc[-2]['high_max'], df.iloc[-1]['Volume'], \
                df_com.iloc[i]['industry'], '60continue'])
        wb.save('trend_follow_60.xlsx')
        print('60일 신고가 지속 발생', df_com.iloc[i]['symbol'])
    elif df.iloc[-2]['Close'] > df.iloc[-3]['high_max_60'] and df.iloc[-1]['Close'] < df.iloc[-2]['high_max_60'] : # 20일 전고점 위 조정
        sheet.append([now, df_com.iloc[i]['market'], df_com.iloc[i]['symbol'], df_com.iloc[i]['code'], \
            df_com.iloc[i]['company_name'], df.iloc[-1]['Close'], df.iloc[-2]['high_max'], df.iloc[-1]['Volume'], \
                df_com.iloc[i]['industry'], '60&adjust'])
        wb.save('trend_follow_60.xlsx')
        print('60일 돌파 후 조정', df_com.iloc[i]['symbol'])
    i += 1   