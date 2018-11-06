import requests
import pandas
import xlrd
def read_tickers(filename): #filename in xlsx presumably
    stock_prices = xlrd.open_workbook(filename)
    sheet = stock_prices.sheet_by_index(0)
    return sheet.col_values(0,1)
print(read_tickers("listcode.xlsm"))

'''
url = 'https://www.alphavantage.co/query'
tickers = ['MSFT','GOOG','X']

for sym in tickers:
	params = dict(
		function='GLOBAL_QUOTE',
		symbol=sym,
		apikey='XXXXXXXXXX'
	)

	resp = requests.get(url=url, params=params)
	data = resp.json()
	print(data["Global Quote"]["05. price"])

	
'''