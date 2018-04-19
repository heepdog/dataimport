#coding:utf-8
# Purpose: swap row/columns of a table
# Created: 28.05.2012
# Copyright (C) 2012, Manfred Moitzi
# License: MIT
from ezodf2.conf import config
from ezodf2 import conf, opendoc, newdoc, Sheet
import ezodf2 as odf
import xlrd
from datetime import datetime
import sys,os

filename2 = 'C:/Python34/chris/test.ods'
filename = 'test2.ods'


def openspreadsheet( filename ):
    
    ext = filename.split('.',1)[1].lower()

    if( ext == 'xls'):
        return xlrd.open_workbook(filename,on_demand=True)
#        getsheet = xlrd.Book.sheet_by_name

    else:
        return odf.opendoc(filename)
#        getsheet = odf.sheets.Sheets._child_by_name

    #mysheet = getsheetref(test,'Summary')
    #print(mysheet.name)

def getsheetref( doc,wb):
    if isinstance(doc, xlrd.Book):
        return doc.sheet_by_name(wb)
    elif isinstance(doc, odf.document.PackagedDocument):
        return doc.sheets[wb]
    else:
        return False

def getcellstring(sheet,cell):

    c,r = cell

    if isinstance(sheet,xlrd.sheet.Sheet):
        cellValue = sheet.cell_value(r,c)
        cellType = sheet.cell_type(r,c)
        if cellType == xlrd.sheet.XL_CELL_ERROR:
            return xlrd.biffh.error_text_from_code[cellValue]

        elif cellType == xlrd.formatting.XL_CELL_DATE:
            if cellValue == 0:
                return ''
            return datetime.strftime(xlrd.xldate.xldate_as_datetime(cellValue,0),"%m/%d/%Y")
        else:
            return str(cellValue)

    elif isinstance(sheet, odf.table.Table):
        cellType = sheet[r,c].value_type
        cellValue = sheet[r,c].value

        if cellType is None:
            return ''
        elif cellType == 'date':
            if int(cellValue.split('-')[0]) <= 1899:
                return ''
            return str(sheet[r,c].display_form)
       
        return str(cellValue)
    else:
        return False


def testezodf2():
    # open spreadsheet document
    # set strategy "all", "all_but_last", "all_less_maxcount'
    config.table_expand_strategy.set_strategy('all_less_maxcount',(3000,150))
    doc = odf.opendoc(filename2)

    print("Spreadsheet contains %d sheets.\n" % len(doc.sheets))
    for sheet in doc.sheets:
        print("Sheet name: '%s'" % sheet.name)
        print("Size of Sheet : (rows=%d, cols=%d)" % (sheet.nrows(), sheet.ncols()) )
        print("-"*40)

    summarysheet = doc.sheets[0]
    #Sheet.
    #Cell.xmlnode
    mycells = summarysheet
    #for cells in mycells:
    #    print(cells.value)
    return doc.sheets[0]


def refsheet():
    NCOLS=10
    NROWS=10

    ods = newdoc('ods','test.ods')

    sheet = Sheet('REFS', size=(NROWS, NCOLS))
    ods.sheets += sheet

    for row in range(NROWS):
        for col in range(NCOLS):
            content = chr(ord('A') + col) + str(row+1)
            sheet[row, col].set_value(content)

    ods.save()

def sumsheet():
    ods = newdoc('ods')

    sheet = Sheet('SUM Formula')
    ods.sheets += sheet

    for col in range(5):
        for row in range(10):
            sheet[row,col].set_value(col*10. + row)

    sheet['F9'].set_value("Summ:")
    sheet['F10'].formula = 'of:=SUM([.A1:.E10])'
    sheet['F1'].formula = 'of:=SUM([.A1];[.B1];[.C1];[.D1];[.E1])'
    ods.saveas('sum_formula.ods')

if __name__ == '__main__':
    testezodf2()

def ignore_exception(exception=Exception, default_val=None):
  """Returns a decorator that ignores an exception raised by the function it
  decorates.

  Using it as a decorator:

    @ignore_exception(ValueError)
    def my_function():
      pass

  Using it as a function wrapper:

    int_try_parse = ignore_exception(ValueError)(int)
  """
  def decorator(function):
    def wrapper(*args, **kwargs):
      try:
        return function(*args, **kwargs)
      except exception:
        return default_val
    return wrapper
  return decorator

def exceptionprint( e ):
   print(
       '''********************************************************************************

Exception - {}
    Caught exception:       {}
             In File:       {}
             On Line:       {}

********************************************************************************'''.format(
       e,
       sys.exc_info()[0].__name__, 
       os.path.basename(sys.exc_info()[2].tb_frame.f_code.co_filename), 
       sys.exc_info()[2].tb_lineno )
         )
