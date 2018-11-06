import requests
import xlrd
import time

def read_tickers(filename): #filename in xlsx presumably
    stock_prices = xlrd.open_workbook(filename)
    sheet = stock_prices.sheet_by_index(0)
    return sheet.col_values(0,1)

url = 'https://www.alphavantage.co/query'
tickers = read_tickers("listcode.xlsm")

prices = {}

for sym in tickers:
	while True:
		params = dict(
			function='GLOBAL_QUOTE',
			symbol=sym,
			apikey='XXXXXXXX'
		)

		resp = requests.get(url=url, params=params)
		data = resp.json()
		if("Global Quote" in data):
			print(str(sym))
			print(data["Global Quote"]["05. price"])
			prices[sym] = data["Global Quote"]["05. price"]
			break
	
		

	
