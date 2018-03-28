#returns string in appropriate format ("mmyy") for 
#the use of other classes in the system.
import datetime
from src import data

def setDate():
	today = datetime.datetime.today()
	data.month = datetime.date.today()
	today = today.replace(day=1)
	data.prevMonth = today - datetime.timedelta(days=1)

		