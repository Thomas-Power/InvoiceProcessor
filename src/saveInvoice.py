'''
Created on 26 Mar 2018

@author: Thomas
'''
import shutil, os, openpyxl
from src import *

def saveInvoice():
    data.invoiceBook.save(data.invoiceBook.strftime('.\\Invoices\\%m-%y' + data.givenAccount))
