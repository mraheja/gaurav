import requests
import openpyxl


target_wb = openpyxl.load_workbook("listcode.xlsx")
target_ws = target_wb.active

tickers = [x.value for x in target_ws["A"]]
print(tickers)

prices = {}
url = 'https://www.alphavantage.co/query'

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


columns = target_ws.max_column + 1
for i,x in enumerate(target_ws["A"],1):
	print(str(i) + " " + str(columns))
	target_ws.cell(row=i, column=columns, value=prices[x.value])
	target_wb.save("listcode.xlsx")


	
