from src import *

def gatherData():
	setOrderBook()
	setInvoiceBook()
	setPreviousBook()
	copyAddress(data.orderSheet, data.prevSheet)
	copyDelivery(data.orderSheet, data.prevSheet)
	gatherOrders()
	copyReference(data.orderSheet, data.prevSheet)