# 지수이동평균(ema) 계산 : 직접 계산과 ewm 하이브러리의 차이 확인 가능 : 결과 다름

from pandas_datareader import data as pdr
import yfinance as yf
yf.pdr_override()
df = pdr.get_data_yahoo('GLPI', start='2021-03-01') 
df['high_max_60'] = df['Close'].rolling(window=60).max()
# df.to_excel("skin.xlsx")
print('전일종가:', df.iloc[-2]['Close'])
print('2일전 최고가 : ', df.iloc[-3]['high_max_60'])
print('금일종가:', df.iloc[-1]['Close'])
print('wjsdlfchlrhrk : ', df.iloc[-2]['high_max_60'])
# 12일 EMA = EMA12
if len(df.Close) < 12:
    print("Stock info is short")

# y = df.Close.values[0]
# m_list = [y]
y = df.iloc[0]['Close']
m_list=[y] 
# 12일 지수이동평균 계산
for k in range(1, len(df.Close)):
    if k < 12:
        a = 2 / (k+1 + 1)
    else:
        a = 2 / (12 + 1)

    y = y*(1-a) + df.Close.values[k]*a
    m_list.append(y)

df.EMA12 = m_list
print(df.EMA12)

# 26일 지수이동평균 계산
y = df.iloc[0]['Close']
m_list=[y] 
for k in range(1, len(df.Close)):
    if k < 26:
        a = 2 / (k+1 + 1)
    else:
        a = 2 / (26 + 1)

    y = y*(1-a) + df.Close.values[k]*a
    m_list.append(y)

df.EMA26 = m_list
print(df.EMA26)

# ewm 라이브러리 이용한 12일, 26일 지수이동평균 산출
df['ema12'] = df['Close'].ewm(12).mean()
df['ema26'] = df['Close'].ewm(26).mean()
# print(df['ema12'])
# print(df['ema26'])