'''Operationally central class, initializes global pointers and 
executes procedures according to their appropriate order.'''
import shutil
import shutil, os
import openpyxl
from src import *

def update():	
	data.holidayBook = openpyxl.load_workbook(".\\code\\holidays.xlsx");
	data.holidaySheet = data.holidayBook.get_sheet_by_name('Sheet1')
	
	setDate()
	produceData()
	produceInvoices()

