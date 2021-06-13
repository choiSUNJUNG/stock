# 일 평균 30만주이상 거래되는 nasdaq, newyork, amex 중 
# 60일 거래량 평균 대비 60일내 일일 최대 거래량이 500% 이상인 종목 중
# 개별종목 체크

from pandas_datareader import data as pdr
import yfinance as yf
yf.pdr_override()
df = pdr.get_data_yahoo('jan', start='2021-01-01') 
df['vol_mean'] = df['Volume'].rolling(window=60).mean()
df['vol_max'] = df['Volume'].rolling(window=60).max()
df.to_excel("each.xlsx")
