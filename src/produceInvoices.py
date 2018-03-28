import shutil, os, datetime
from src import *

def produceData():
	for data.givenAccount in os.listdir(data.prevMonth.strftime('.\\Invoices\\%m-%Y Invoices')):
		makeInvoice()
		saveInvoice()
