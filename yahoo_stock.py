from pandas_datareader import data as pdr
import yfinance as yf
import pandas as pd
import matplotlib.pyplot as plt
import openpyxl
import datetime
import time

yf.pdr_override()
wb = openpyxl.Workbook()
# wb.save('watch_data.xlsx')
sheet = wb.active
# cell name : date, simbol, company name, upper or lower or narrow band 
# sheet.append(['time', 'simbol', 'company_name', 'bollinger_band'])
# wb.save('watch_data.xlsx')
#회사 데이터 읽기
df_com = pd.read_excel("usa_company.xlsx")
# print(df_com)
# len(df_com)
# print(len(df_com))

while True:
    try:
        time.sleep(10)
        i = 1
        for i in range(len(df_com)):
            now = datetime.datetime.now()
            df = pdr.get_data_yahoo(df_com.iloc[i]['simbol'], start = '2020-01-01')
            df['MA20'] = df['Close'].rolling(window=20).mean()
            df['stddev'] = df['Close'].rolling(window=20).std()
            df['upper'] = df['MA20'] + (df['stddev']*3)
            df['lower'] = df['MA20'] - (df['stddev']*3)
            df['vol_avr'] = df['Volume'].rolling(window=5).mean()
            df['bandwidth'] = (df['upper'] - df['lower']) / df['MA20'] * 100
            df = df[19:]
            cur_price = df.iloc[-1]['Close']
            df_u = df.iloc[-1]['upper']
            df_l = df.iloc[-1]['lower']
            df_b = df.iloc[-1]['bandwidth']
            df_v = df.iloc[-2]['vol_avr']
            print(df_com.iloc[i]['simbol'])
            print('볼린저 : ',df.iloc[-1]['MA20'], df.iloc[-1]['upper'], df.iloc[-1]['lower'])
            print('볼린저 밴드폭 : ',df.iloc[-1]['bandwidth'], '%')


            if (cur_price >= df_u or cur_price <= df_l) & (df_v > 300000) :
                print('watch1', df_com.iloc[i]['simbol'])
                sheet.append([now, df_com.iloc[i]['simbol'], df_com.iloc[i]['company_name'], 'upper/lower'])
                wb.save('watch_data.xlsx')
            elif df_b <= 20:
                sheet.append([now, df_com.iloc[i]['simbol'], df_com.iloc[i]['company_name'], 'narrow'])
                wb.save('watch_data.xlsx')
            i += 1   
            time.sleep(1)
    except Exception as e:
        print(e)
        time.sleep(1)
