import datetime

def formatDay(date, curMonth):
		day = date[:2]
		cDay = datetime.date(curMonth.year, curMonth.month, day)
		return(cDay.isoweekday())