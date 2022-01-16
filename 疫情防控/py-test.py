import pandas as pd
import numpy as np
path = 'D:/Users/user/Desktop/填表记录_员工疫情防控期间行动轨迹情况报告基础表.xls'
df=pd.read_excel(path,sheet_name='填表记录')#这个会直接默认读取到这个Excel的第一个表单
data=df.values

#data[0][10] = 'add New Values'

print("获取到所有的值:\n{0}".format(data))#格式化输出
print(len(data)) #行数
print(data.shape) #行列号
print(data.ndim) #维数

details=[]
prefixrowIndex = 0
rowIndex = 0
prefixrowName = ''
rowName=''

rowName = data[rowIndex][1]
print(rowName)

for i in range(0,len(data)):
    details.append(data[i][5]+','+data[i][6]+','+data[i][7])

detailsRange=[]

for rowIndex in range(0,len(data)):
    nxtIndexValue = -1
    currentName = data[rowIndex][1]#姓名
    currentDate = data[rowIndex][2]#日期
    currentTime = data[rowIndex][3]#时间
    currentStr = details[rowIndex]
    for nxtIndex in range(rowIndex,len(data)):
        if data[nxtIndex][1]==currentName and data[nxtIndex][2]==currentDate and data[nxtIndex][3]==currentTime:
                currentStr = currentStr+',' + '换行' + ',' +details[nxtIndex]
                nxtIndexValue = nxtIndex

    detailsRange.append(currentStr)
    if nxtIndexValue>-1:
        rowIndex = nxtIndexValue

print(detailsRange)
        
    

