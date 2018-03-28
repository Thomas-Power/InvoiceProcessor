from src import *
import datetime, calendar 

class makeInvoice():


	global curDate 
	global curDay
	
	'''Used to control active row'''
	global gap
	global curRow
	gap = 0
	curRow = 0
	'''used to control active day of week'''
	global cov
	cov = 0
	
	def makeInvoice(self):
		setOrderBook()
		setInvoiceBook()
		copyAddress(data.orderSheet, data.invoiceSheet)
		copyDelivery(data.invoiceSheet)
		copyReference(data.invoiceSheet)
		for i in calendar.monthrange(data.month.year, data.month.month):
			curRow = i - gap
			curDate = i + cov
			curDay = weekDay(curDate)
			checkDate()
			if(curDay == 5):
				cov = 0
