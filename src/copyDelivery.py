import shutil, os, openpyxl
from src import *

def copyDelivery(target, source):
	if(target == data.orderSheet):
		target.cell(row=22, column=5).value = source.cell(row=50, column=5).value
	elif(source == data.orderSheet):
		target.cell(row=50, column=5).value = source.cell(row=22, column=5).value
	else:
		target.cell(row=50, column=5).value = source.cell(row=50, column=5).value