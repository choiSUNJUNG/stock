from pandas_datareader import data as pdr
import yfinance as yf
import matplotlib.pyplot as plt

yf.pdr_override()
sec = pdr.get_data_yahoo('005930.KS', start='2018-05-04')
msft = pdr.get_data_yahoo('MSFT', start='2018-05-04')
soxl = pdr.get_data_yahoo('SOXL', start='2018-05-04')

print(sec.head(10))
print(msft.tail(10))
print('SOXL: ', soxl.tail(10))

# plt.plot(sec.index, sec.Close, 'b', label='삼성전자')
# plt.plot(msft.index, msft.Close, 'b', label='마이크로소프트')
plt.plot(soxl.index, soxl.Close, 'r--', label='SOXL_ETF')
plt.legend(loc='best')
plt.show()