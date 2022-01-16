from asyncio.windows_events import NULL
import os
from turtle import end_fill
import pandas as pd
import xlrd

#参考文档
#https://zhuanlan.zhihu.com/p/56808884

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

class OSDirectory:
    DirectoryList = set()
    FileList = set()
    def ReadDirectory(self,directory):
        for root,dirs,files in os.walk(directory): 
            for dir in dirs: 
                #self.DirectoryList.add(os.path.join(root,dir).decode('gbk').encode('utf-8'))
                self.DirectoryList.add(dir)
            for file in files: 
                #self.FileList.add(os.path.join(root,file).decode('gbk').encode('utf-8'))
                self.FileList.add(file)
            return self.DirectoryList,self.FileList

#excel文件类
class ExcelFile:    
    #Excel工作簿
    Book = NULL
    CurrentSheet = NULL
    CurrentCell = NULL

    #打开文件函数
    def __open_file(self,strPath):
        ret = False
        ret = os.path.exists(strPath)
        if(ret):
            #文件存在
            print(ret)    
        return ret,strPath
    #打开Excel文件，赋值给excelFileBook
    def open_excel_file(self,strFileName):
        global Book
        retValue = self.__open_file(strFileName)
        if(retValue[0]):
            #文件存在
            strFileName = retValue[1]
            #打开excel文件
            self.Book = xlrd.open_workbook(strFileName)

    #打开Excel文件中的Sheet，需要给出Sheet名称
    def open_sheet(self,sheetName):
        # 获取第1个表单
        self.CurrentSheet = self.Book.sheet_by_name(sheetName)

    #打开Excel文件中的Sheet中的单元格，需要Sheet名称，行列序号。
    def open_cell(self,sheetName,col,row):
        self.open_sheet(sheetName)
        self.CurrentCell = self.CurrentSheet.cell_value(col,row)

#打开excel文件
excelFile  = ExcelFile()
excelFile.open_excel_file("d:\学员信息导入.xls")
excelFile.open_sheet("Sheet1")
excelFile.open_cell("Sheet1",0,0)

print ("表单数量:", excelFile.Book.nsheets)
print ("表单名称:", excelFile.Book.sheet_names())
print (u"表单 %s 共 %d 行 %d 列" % (excelFile.CurrentSheet.name, excelFile.CurrentSheet.nrows,excelFile.CurrentSheet.ncols))
print ("第0行第0列:", excelFile.CurrentCell)

for sheetItem in excelFile.Book.sheets():
    for r in range(sheetItem.nrows):
        print ("行内容：" ,sheetItem.row(r))

osDirectory = OSDirectory()
print ("目录 = " , osDirectory.ReadDirectory("d:\\tools\\"))
# 输出指定行
#df = pd.read_excel("")
#读取excel文件内容
#data=df.head(2)

