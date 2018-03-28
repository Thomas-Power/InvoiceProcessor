'''
Created on 26 Mar 2018

@author: Thomas
'''

from src import *

def checkDate():
    if(makeInvoice.curDay != None):
        for i in range (0, 5, 1):
            if(makeInvoice.curDay == data.orderSheet.cell((14 + i), 1).value):
                if(isHoliday() != False):
                    pushWeek()
                else:
                    postData()
        else:
            makeInvoice.gap += 1
    else:
        makeInvoice.gap += 1