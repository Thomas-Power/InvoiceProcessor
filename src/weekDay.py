'''
Created on 26 Mar 2018

@author: Thomas
'''
import copy
from src import *

def weekDay():
    def weekDay(self, date):
        temp = data.month.copy()
        temp.replace(day=date)
        return temp.isoweekday()