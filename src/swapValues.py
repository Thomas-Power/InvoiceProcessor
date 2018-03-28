'''
Created on 26 Mar 2018
Swaps the product values of current day with previous or next day
@author: Thomas
'''
from src import *

def swapValues(x):
    data.invoiceSheet.cell(row=(makeInvoice.curRow + x), column=2).value = data.orderSheet.cell(row=makeInvoice.curDay, column=2).value
    data.invoiceSheet.cell(row=(makeInvoice.curRow + x), column=5).value = data.orderSheet.cell(row=makeInvoice.curDay, column=5).value
    if(x == 0):
        data.invoiceSheet.cell(row=(makeInvoice.curRow + x), column=2).value = formatDate(makeInvoice.curDay+1)