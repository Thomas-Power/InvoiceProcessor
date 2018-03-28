import shutil, os, openpyxl
from src import *

def setPreviousBook():
	data.previousBook = openpyxl.load_workbook(data.prevMonth.strftime('.\\Accounts\\ + %m-%Y + Invoices\\' + data.givenAccount))
	data.previousSheet = data.previousBook.get_sheet_by_name('Sheet1')
