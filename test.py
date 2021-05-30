from pandas_datareader import data as pdr
from pandas import DataFrame
import yfinance as yf
import pandas as pd
import openpyxl
import datetime
import time
import math

yf.pdr_override()
# wb = openpyxl.Workbook()
# cell name : date, simbol, company name, upper or lower or narrow band 
#회사 데이터 읽기

now = datetime.datetime.now()
print("현재시간 : ", now)
df = pdr.get_data_yahoo('MSFT', start='2018-12-25')  # 기간 70일(60일 이동평균까지 구할 수 있도록)
# df = pdr.get_data_yahoo('MLND', period = '70d')  # 기간 70일(60일 이동평균까지 구할 수 있도록)
df['ema12'] = df['Close'].ewm(12).mean()    # 지수이동평균(12일)
df['ema26'] = df['Close'].ewm(26).mean()    # 지수이동평균(26일)
df['ma15'] = df['Close'].rolling(window=15).mean() # 단순이동평균(SMA 15일)
df['ma20'] = df['Close'].rolling(window=20).mean()
df['ma60'] = df['Close'].rolling(window=60).mean()
# ADX 구하기
df['pdm'] = abs(df['High'] - df['High'].shift(1))
df['mdm'] = abs(df['Low'].shift(1) - df['Low'])
df['cl'] = abs(df['Close'] - df['Close'].shift(1))
# df = df.dropna(axis = 0)
# df['tr'] = max(df['pdm'], df['mdm'], df['cl'])
a = max(1, 2, 3)
df['pdm11'] = df['pdm'].rolling(window=11).mean()
df['mdm11'] = df['mdm'].rolling(window=11).mean() 
# ADX 산출위한 TR(True Range(실변동폭)) 구하기
# 1. 오늘의 고가와 저가 차이(tr1) : 절대값
# 2. 어제의 종가와 오늘의 고가 차이(tr2) : 절대값
# 3. 어제의 종가와 오늘의 저가 차이(tr3) : 절대값 
# 이들 3개값중 가장 큰 값이 실변동폭
df['tr1'] = abs(df['High'] - df['Low'])
df['tr2'] = abs(df['High'] - df.iloc[-2]['Close'])
df['tr3'] = abs(df['Low'] - df.iloc[-2]['Close'])
df['tr'] = max(df['tr1'], df['tr2'], df['tr3'])
print(df['tr1'], df['tr2'], df['tr3'])
# if p1 == 1:
#     df['tr'] = df['pdm']
# elif p2 == True:
#     df['tr'] = df['mdm']
# else:
#     df['tr'] = df['cl']


# i = 1
# for i in range(len(df)):
#     if (df.iloc(i)['pdm'] > df.iloc(i)['mdm']) & (df.iloc(i)['pdm'] > df.iloc(i)['cl']) :
#         df['tr'] = df['pdm']
#     elif df.iloc(i)['mdm'] > df.iloc(i)['pdm'] and df.iloc(i)['mdm'] > df.iloc(i)['cl'] :
#         df['tr'] = df['mdm']
#     else:
#         df['tr'] = df['cl']
#     i += 1

df.to_excel("dd11.xlsx")



# if (pdm100 - mdm100) > 0 :
#     if (pdm100 - cl100) > 0:
#         df['tr'] = df['pdm']
# # # # elif df['mdm'] > df['pdm'] & df['mdm'] > df['cl']:
# # # #     df['tr'] = df['mdm']
# else:
#     df['tr'] = df['cl']

# df['tr_11'] = df['tr'].rolling(window=11).mean() 


# df['pdi'] = df['pdm_11'] / df['tr_11']
# df['mdi'] = df['mdm_11'] / df['tr_11']
# df['dx'] = (df['pdi'] - df['mdi']) / (df['pdi'] + df['mdi'] * 100)
# df['adx'] = df['dx'] / 11
# df = df.dropna(axis = 0)    # DATA table에서 NaN 있는 행 삭제하기
# df.to_excel("dd11.xlsx")
# # 지수 이동 평균 구하기
# 

# """stochastics 구하기"""
# df['min_14'] = df['Low'].rolling(window=14).min()
# df['max_14'] = df['High'].rolling(window=14).max()
# df.min = df['min_14']
# df.max = df['max_14']
# df['sto_K_14'] = (df['Close'] - df.min) / (df.max - df.min) * 100   #stochastics_fast의 %K
# df['sto_D_5'] = df['sto_K_14'].rolling(5).mean()    #stochastics_slow의 %K
# df['sto_DS_3'] = df['sto_D_5'].rolling(3).mean()    #stochastics_slow의 %D
# # """MACD & signal 구하기 """
# df['MACD'] = df['ema12'] - df['ema26']
# df['signal'] = df['MACD'].rolling(window=9).mean()
# # print("오늘 종가 : ",df.iloc[-1]['Close'])    
# print(df.iloc[-2]['ema12'], df.iloc[-2]['ema26'], df.iloc[-2]['MACD'])
# print(df.iloc[-1]['ema12'], df.iloc[-1]['ema26'], df.iloc[-1]['MACD'])
# print(df.iloc[-1]['sto_D_5'])
# print(df.iloc[-1]['adx'])
# df.to_excel("dd11.xlsx")
