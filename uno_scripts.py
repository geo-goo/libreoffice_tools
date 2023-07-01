#!/bin/python3

import uno

def my_first_macro_calc():
    doc = XSCRIPTCONTEXT.getDocument()
    try :
        cell = doc.Sheets['test']['A1']  # com.sun.star.sheet.XSpreadsheetDocument
        cell.setString('Hello World in Python in Calc')
    except :
        cell = doc.Sheets[0]['A1']
        cell.setString('Hello World in Python in Calc')
    return
