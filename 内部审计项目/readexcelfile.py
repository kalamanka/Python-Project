from asyncio.windows_events import NULL
import os
from turtle import end_fill
import pandas as pd
import xlrd
# 打开 xls 文件
#book = xlrd.open_workbook("test.xls")
#print "表单数量:", book.nsheets
#print "表单名称:", book.sheet_names()
# 获取第1个表单
#sh = book.sheet_by_index(0)
#print u"表单 %s 共 %d 行 %d 列" % (sh.name, sh.nrows,sh.ncols)
#print "第二行第三列:", sh.cell_value(1, 2)
# 遍历所有表单
#for s in book.sheets():
#for r in range(s.nrows):
# 输出指定行
#print s.row(r)

#excel文件类
class ExcelFile(object):    
    ExcelFileBook = NULL
    #打开文件函数
    def open_file(strPath):
        ret = False
        ret = os.path.exists(strPath)
        if(ret):
            #文件存在
            print(ret)    
        return ret,strPath
    #打开Excel文件，赋值给excelFileBook
    def open_excel_file(strFileName):
        ret = open_file(strFileName)
        if(ret[0]):
            #文件存在
            strFileName = ret[1]
            #打开excel文件
            ExcelFileBook = xlrd.open_workbook(strFileName)


#打开excel文件
excelFile  = ExcelFile()
excelFile.open_excel_file("d:\学员信息导入.xls")

#df = pd.read_excel("")
#读取excel文件内容
#data=df.head(2)

