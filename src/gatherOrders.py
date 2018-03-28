from src import *

def copyOrders():
	days = {False, False, False, False, False}
	
	for i in range (14, 50, 1):
		currentDay = formatDate(data.invoiceSheet.cell(row=i, column=1).value, data.prevMonth)
		for e in range (0, 5, 1):
			if(currentDay == e):
				'''bool must be checked to ensure exceptional dates are not copied'''
				if(days[e] == True):
					data.orderSheet.cell(row=(14 + e), column=2).value = data.invoiceSheet.cell(row=i, column=2).value
					data.orderSheet.cell(row=(14 + e), column=5).value = data.invoiceSheet.cell(row=i, column=5).value
				days[e] = True