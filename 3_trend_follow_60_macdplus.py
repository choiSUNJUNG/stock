# 리처드 데니스의 추세추종전략 : 12주(60일선) 전고점 돌파 종목 추출 + MACD 0 이상 

from pandas_datareader import data as pdr
import yfinance as yf
import pandas as pd
import openpyxl
import datetime
import time

def macd(stick, period):
    df = pdr.get_data_yahoo(stick, start=period) 

    # 12일 EMA = EMA12
    if len(df.Close) < 26:
        print("Stock info is short")

    # y = df.Close.values[0]
    # m_list = [y]
    y1 = df.iloc[0]['Close']
    y2 = df.iloc[0]['Close']
    m_list1 = [y1] 
    m_list2 = [y2]
    # 12일 지수이동평균 계산
    for i in range(len(df)):
        if i < 12:
            a12 = 2 / (i+1 + 1)
        else:
            a12 = 2 / (12 + 1)
        y1 = y1*(1-a12) + df.Close.values[i]*a12
        m_list1.append(y1)

    # 26일 지수이동평균 계산    
    for k in range(len(df)):
        if k < 26:
            a26 = 2 / (k+1 + 1)
        else:
            a26 = 2 / (26 + 1)
        y2 = y2*(1-a26) + df.Close.values[k]*a26
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
sheet.append(['time', 'market', 'symbol', 'code', 'company_name', 'price', 'pre_high_low', 'volume', 'industry', 'macd', 'sign'])
wb.save('trend_follow_60_macdplus.xlsx')

#회사 데이터 읽기
# df_com = pd.read_excel("300k_day_coms.xlsx")
df_com = pd.read_excel("300k_day_coms.xlsx")
now = datetime.datetime.now()
start_day = '2021-01-01'
i = 1
for i in range(len(df_com)):
    # df = pdr.get_data_yahoo(df_com1.iloc[i]['Symbol'], period = '1mo')  # 기간 1month
    df = pdr.get_data_yahoo(df_com.iloc[i]['symbol'], period = '70d')  # 기간 10일
    df['high_max'] = df['Close'].rolling(window=20).max()
    df['high_max_60'] = df['Close'].rolling(window=60).max()
    # df['low_min'] = df['Close'].rolling(window=10).min()
    # print(df)

    if df.iloc[-2]['Close'] < df.iloc[-3]['high_max_60'] and df.iloc[-1]['Close'] > df.iloc[-2]['high_max_60']: 
        # 60일 전고점 돌파 첫날
        macd(df_com.iloc[i]['symbol'], start_day)
        if macd(df_com.iloc[i]['symbol'], start_day) > 0 :
            print(macd(df_com.iloc[i]['symbol'], '2021-01-01'))
            sheet.append([now, df_com.iloc[i]['market'], df_com.iloc[i]['symbol'], df_com.iloc[i]['code'], \
                df_com.iloc[i]['company_name'], df.iloc[-1]['Close'], df.iloc[-2]['high_max_60'], df.iloc[-1]['Volume'], \
                    df_com.iloc[i]['industry'], macd(df_com.iloc[i]['symbol'], start_day), '●60overday'])
            wb.save('trend_follow_60_macdplus.xlsx')
            print('60일 전고점 돌파 첫날', df_com.iloc[i]['symbol'])
    elif df.iloc[-2]['Close'] > df.iloc[-3]['high_max_60'] and df.iloc[-1]['Close'] > df.iloc[-2]['high_max_60'] : # 60일 전고점 위 상승 지속
        macd(df_com.iloc[i]['symbol'], start_day)
        if macd(df_com.iloc[i]['symbol'], start_day) > 0 :
            sheet.append([now, df_com.iloc[i]['market'], df_com.iloc[i]['symbol'], df_com.iloc[i]['code'], \
                df_com.iloc[i]['company_name'], df.iloc[-1]['Close'], df.iloc[-2]['high_max_60'], df.iloc[-1]['Volume'], \
                    df_com.iloc[i]['industry'], macd(df_com.iloc[i]['symbol'], start_day), '60continue'])
            wb.save('trend_follow_60_macdplus.xlsx')
            print('60일 신고가 지속 발생', df_com.iloc[i]['symbol'])
    elif df.iloc[-2]['Close'] > df.iloc[-3]['high_max_60'] and df.iloc[-1]['Close'] < df.iloc[-2]['high_max_60'] : # 20일 전고점 위 조정
        macd(df_com.iloc[i]['symbol'], start_day)
        if macd(df_com.iloc[i]['symbol'], start_day) > 0 :
            sheet.append([now, df_com.iloc[i]['market'], df_com.iloc[i]['symbol'], df_com.iloc[i]['code'], \
                df_com.iloc[i]['company_name'], df.iloc[-1]['Close'], df.iloc[-2]['high_max_60'], df.iloc[-1]['Volume'], \
                    df_com.iloc[i]['industry'], macd(df_com.iloc[i]['symbol'], start_day), '60&adjust'])
            wb.save('trend_follow_60_macdplus.xlsx')
            print('60일 돌파 후 조정', df_com.iloc[i]['symbol'])
    i += 1   