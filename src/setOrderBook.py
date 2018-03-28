import shutil, os, openpyxl
from src import *

def setOrderBook():
	if(os.path.isdir('.\\Order Sheets\\' + data.givenAccount)):
		data.orderBook = openpyxl.load_workbook('.\\Order Sheet\\' + data.givenAccount)
		data.orderSheet = data.orderBook.get_sheet_by_name('Sheet1')
	else:
		data.orderBook = openpyxl.load_workbook('.\\Code\\orderTemplate.xlxs')
		data.orderSheet = data.orderBook.get_sheet_by_name('Sheet1')