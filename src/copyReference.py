'''Copy reference number and previous iterations'''

import shutil, os, openpyxl
from src import *

def copyReference(target, source):
	if(target == data.orderSheet):
		target.cell(row=4, column=5).value = source.cell(row=4, column=5).value
		target.cell(row=5, column=5).value = source.cell(row=4, column=5).value
	else:
		target.cell(row=4, column=5).value = source.cell(row=4, column=5).value