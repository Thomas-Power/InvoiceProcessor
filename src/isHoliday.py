'''
Created on 26 Mar 2018

@author: Thomas
'''

from src import *

def isHoliday():
    check = False
    i = 1
    while(check == False):
        if(makeInvoice.curDate == data.holidaySheet.cell(1, i).value):
            check == True
        if(makeInvoice.curDate > data.holidaySheet.cell(1, i).value):
            break
    return check