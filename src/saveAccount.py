'''
Created on 26 Mar 2018

@author: Thomas
'''
import shutil, os, openpyxl
from src import *

def saveAccount():
    data.orderBook.save(data.orderBook.strftime('.\\Account Sheets\\' + data.givenAccount))
