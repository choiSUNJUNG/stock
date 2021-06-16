# 120일 이동평균 상승 & stochastic 20이하 종목 추출 

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
sheet.append(['time', 'market', 'symbol', 'code', 'company_name', 'industry', 'stochastic'])
wb.save('7_120_stoch.xlsx')

#회사 데이터 읽기
df_com = pd.read_excel("step2_300k_day_coms.xlsx")
now = datetime.datetime.now()

i = 1
for i in range(len(df_com)):
    # df = pdr.get_data_yahoo(df_com1.iloc[i]['Symbol'], period = '1mo')  # 기간 1month
    df = pdr.get_data_yahoo(df_com.iloc[i]['symbol'], period = '6mo')  # 기간 10일
    df['min_14'] = df['Low'].rolling(14).min()
    df['max_14'] = df['High'].rolling(14).max()
    df.min = df['min_14']
    df.max = df['max_14']
    df['sto_K_14'] = (df['Close'] - df.min) / (df.max - df.min) * 100   #stochastics_fast의 %K
    df['sto_D_5'] = df['sto_K_14'].rolling(5).mean()    #stochastics_slow의 %K
    df['sto_DS_3'] = df['sto_D_5'].rolling(3).mean()    #stochastics_slow의 %D
    # df['ma20'] = df['Close'].rolling(window=20).mean()
    # df['ma60'] = df['Close'].rolling(window=60).mean()
    df['ma120'] = df['Close'].rolling(window=120).mean()
    # print(df.iloc[-1]['sto_D_5'])
    # print(df.iloc[-1]['ma120'], df.iloc[-2]['ma120'], df.iloc[-3]['ma120'])
    if df.iloc[-1]['ma120'] >= df.iloc[-2]['ma120'] >= df.iloc[-3]['ma120'] and df.iloc[-1]['sto_D_5'] <= 50 and df.iloc[-1]['sto_D_5'] > df.iloc[-1]['sto_DS_3'] :
        sheet.append([now, df_com.iloc[i]['market'], df_com.iloc[i]['symbol'], df_com.iloc[i]['code'], \
            df_com.iloc[i]['company_name'], df.iloc[-1]['Close'], df_com.iloc[i]['industry'], df.iloc[-1]['sto_D_5']])
        wb.save('7_120_stoch.xlsx')
        print('120 상승, 스토캐스틱 20이하 & %K > %S', df_com.iloc[i]['symbol'])
    
    i += 1   
    