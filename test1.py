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
# df = pdr.get_data_yahoo('AAPL', start='2021-01-01', end='2021-06-30')  # 기간 70일(60일 이동평균까지 구할 수 있도록)
df = pdr.get_data_yahoo('BSPE', period = '10d')
# df = pdr.get_data_yahoo('MLND', period = '70d')  # 기간 70일(60일 이동평균까지 구할 수 있도록)

df.to_excel('dd120.xlsx')