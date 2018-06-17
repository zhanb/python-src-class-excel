#!/usr/bin/env python
#coding=gbk
###Auth:zhangbo_billing
###Date:20180110
###Desc:excel简单操作类
import xlrd

from xlutils.copy import copy
class Excel:
    #excel简单操作类
    def __init__(self):
        self.handle,self.mhandle= None, None
        self.name_sheets = {}
        self.newsheet = None
        self.filename = None
        self.ceil_list = []
        self.rows_list = []
    def setSheet(self,fileName):
        self.filename = fileName
        self.handle = xlrd.open_workbook(fileName)
        self.sheets = dict([(sheetname,self.handle.sheet_by_name(sheetname)) for sheetname in self.handle.sheet_names()])
        self.mhandle = copy(xlrd.open_workbook(self.filename, formatting_info=True))
    def getCeilList(self,sheetName,clos):
        for i in range(self.sheets[sheetName].nrows):
            self.ceil_list.append(self.sheets[sheetName].cell(i,clos).value)
        return self.ceil_list
    def getRowsList(self,sheetName):
        for i in range(self.sheets[sheetName].nrows):
            self.rows_list.append(self.sheets[sheetName].row_values(i))
        return self.rows_list
    def getNameByIndex(self,index):
        return self.handle.sheet_by_index(index).name
    def modifyToNewSheet(self,row,coil,str,sheet = 0):
        self.newsheet = self.mhandle.get_sheet(sheet)
        self.newsheet.write(row,coil,str)
    def modifySave(self,newFileName):
        self.mhandle.save(newFileName)
