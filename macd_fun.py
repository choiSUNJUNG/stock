#MACD 함수처리
from pandas_datareader import data as pdr
import yfinance as yf
import pandas as pd
import openpyxl
import datetime
import time
import math

wb = openpyxl.Workbook()

yf.pdr_override()

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

macd('HOFV', '2021-01-01') 
print(macd('HOFV', '2021-01-01'))


