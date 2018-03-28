import shutil, os, datetime
from src import *
class produceData():
	for givenAccount in os.listdir(data.prevMonth.strftime('.\\%m-%y Invoices')):
		gatherData()
		saveAccount()
