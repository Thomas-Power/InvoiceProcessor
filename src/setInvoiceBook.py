import shutil, os, openpyxl
from src import *

def setInvoiceBook():
	if(os.path.isdir(data.month.strftime('.\\Accounts\\ + %m-%Y + Invoices\\' + data.givenAccount))):
		data.invoiceBook = openpyxl.load_workbook(data.month.strftime('.\\Accounts\\ + %m-%Y + Invoices\\' + data.givenAccount))
		data.invoiceSheet = data.invoiceBook.get_sheet_by_name('Sheet1')
	else:
		data.invoiceBook = openpyxl.load_workbook('.\\Code\\invoiceTemplate.xlxs')
		data.invoiceSheet = data.invoiceBook.get_sheet_by_name('Sheet1')