#!/usr/bin/env python
#coding=gbk

from sys import path
path.append('../class')
import gc
import deal_excel as excel
if __name__ == "__main__":
    myExcel = excel.Excel()
    myExcel.setSheet("../data/数据迁移比对.xls")
    name = myExcel.getNameByIndex(0)
    print(name.encode('gb2312'))
    ceil=myExcel.getCeilList(name,2)
    for i in ceil:
        print (i.encode('gb2312'))
        rows = myExcel.getRowsList(name)
    for i in rows:
        print(i[0].encode('gb2312'))
        gc.collect()
