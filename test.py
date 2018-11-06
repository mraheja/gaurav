import xlrd

filename = "listcode.xlsm"
stock_prices = xlrd.open_workbook(filename)
sheet = stock_prices.sheet_by_index(0)

i = 1
max = 1
while (not (len(sheet.cell(i,0).value == 0)):
	j = 1
	while(not(len(sheet.cell(i,j).value == 0)):
		if(j > max):
			max = j

	