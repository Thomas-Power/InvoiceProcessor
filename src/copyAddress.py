import shutil, os, openpyxl
from src import *

def copyAddress(target, source):
	for i in range(6, 13, 1):
		target.cell(row=i, column=2).value = source.cell(row=i, column=2).value
		
