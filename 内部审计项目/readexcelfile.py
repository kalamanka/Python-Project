
#from asyncio.windows_events import NULL
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
        #global Book
        retValue = self.__open_file(strFileName)
        if(retValue[0]):
            #文件存在
            strFileName = retValue[1]
            #打开excel文件
            self.Book = xlrd.open_workbook(strFileName)

    #打开Excel文件中的Sheet，需要给出Sheet名称
    def open_sheet(self,sheetName):
        # 获取第1个表单
        #global CurrentSheet
        self.CurrentSheet = self.Book.sheet_by_name(sheetName)

    #打开Excel文件中的Sheet中的单元格，需要Sheet名称，行列序号。
    def open_cell(self,sheetName,col,row):
        #global CurrentCell
        self.open_sheet(sheetName)
        self.CurrentCell = self.CurrentSheet.cell_value(col,row)

#打开excel文件
excelFile = ExcelFile()
#excelFile.open_excel_file("d:\学员信息导入.xls")
excelFile.open_excel_file("/home/huangjw/Desktop/2022内审汇总需求/内审统01表_单位.xls")
excelFile.open_sheet("内审统01表")


print ("表单数量:", excelFile.Book.nsheets)
print ("表单名称:", excelFile.Book.sheet_names())
print (u"表单 %s 共 %d 行 %d 列" % (excelFile.CurrentSheet.name, excelFile.CurrentSheet.nrows,excelFile.CurrentSheet.ncols))

#读取机构类型
excelFile.open_cell("内审统01表",12,1)
if(excelFile.CurrentCell == "机构类型"):
    #读取机构类型
    excelFile.open_cell("内审统01表",12,3)
    strOrgnization = excelFile.CurrentCell.strip()[0:2]
    #打印机构编码
    print ("机构类型：" ,strOrgnization)

#读取总审计师
excelFile.open_cell("内审统01表",40,1)
if(excelFile.CurrentCell == "是否设置总审计师"):
    #读取总审计师
    excelFile.open_cell("内审统01表",40,4)
    strAuditMaster = excelFile.CurrentCell.strip()[0:1]
    #打印读取总审计师
    print ("是否设置总审计师" ,strAuditMaster)

#读取是否设置内部审计机构		
excelFile.open_cell("内审统01表",47,1)
if(excelFile.CurrentCell == "是否设置内部审计机构"):
    #读取是否设置内部审计机构
    excelFile.open_cell("内审统01表",47,4)
    strInnerAuditOffice = excelFile.CurrentCell.strip()[0:1]
    #打印读取是否设置内部审计机构
    print ("是否设置内部审计机构" ,strInnerAuditOffice)

#读取是否设置内部审计机构		
excelFile.open_cell("内审统01表",54,1)
if(excelFile.CurrentCell == "是否独立设置内部审计机构"):
    #读取是否设置内部审计机构
    excelFile.open_cell("内审统01表",54,4)
    strIndependentInnerAuditOffice = excelFile.CurrentCell.strip()[0:1]
    #打印读取是否设置内部审计机构
    print ("是否独立设置内部审计机构" ,strIndependentInnerAuditOffice)

#读取实有人员数		
excelFile.open_cell("内审统01表",60,5)
if(excelFile.CurrentCell == "实有人员数"):
    nPersonNumber = 0
    #读取实有人员数	
    excelFile.open_cell("内审统01表",60,6)
    strPersonNumber = excelFile.CurrentCell.strip()
    if(strPersonNumber.length() > 0):
        nPersonNumber = strPersonNumber[0:1]
    #打印实有人员数	
    print ("实有人员数" ,strIndependentInnerAuditOffice)

#读取实有人员数		
excelFile.open_cell("内审统01表",60,7)
if(excelFile.CurrentCell == "其中：专职人员（人）"):
    nFulltimePersonNumber = 0
    #读取实有人员数	
    excelFile.open_cell("内审统01表",60,8)
    strPersonNumber = excelFile.CurrentCell.strip()
    if(strPersonNumber.length() > 0):
        nFulltimePersonNumber = strPersonNumber[0:1]
    #打印实有人员数	
    print ("其中：专职人员（人）" ,strIndependentInnerAuditOffice)

#读取实有人员数		
    nFulltimePersonNumber = 0
    #读取实有人员数	
    excelFile.open_cell("内审统01表",60,8)
    strPersonNumber = excelFile.CurrentCell.strip()
    if(strPersonNumber.length() > 0):
        nFulltimePersonNumber = strPersonNumber[0:1]
    #打印实有人员数	
    print ("其中：专职人员（人）" ,strIndependentInnerAuditOffice)

for sheetItem in excelFile.Book.sheets():
    for r in range(sheetItem.nrows):
        print ("行内容：" ,sheetItem.row(r))

osDirectory = OSDirectory()
#print ("目录 = " , osDirectory.ReadDirectory("d:\\tools\\"))
# 输出指定行
#df = pd.read_excel("")
#读取excel文件内容
#data=df.head(2)

