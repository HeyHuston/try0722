import xlwt
import xlrd
import xlutils
from xlutils.copy import copy
import os
import sys

workbook1=xlrd.open_workbook(r'C:\Users\33590\Desktop\GIS整合\0329GIS2.xlsx') 
sheet1=workbook1.sheet_by_name('合')  
nrows1=sheet1.nrows  
ncols1=sheet1.ncols  
sheet2=workbook1.sheet_by_name('附表')  
nrows2=sheet2.nrows  
ncols2=sheet2.ncols  

xls=copy(wb=workbook1)
sheet=xls.get_sheet(4)

print(nrows1)
print(nrows2)

#sys.exit()
m=0
for row1 in range(1,nrows1):
    flag=0
    menpaihao1=sheet1.cell(row1,3).value  
    if menpaihao1=='':
        continue
    else:
        menpaihao1=sheet1.cell(row1,4).value  
        for row2 in range(1,nrows2):
            menpaihao2=sheet2.cell(row2,1).value
            if menpaihao2==menpaihao1:   
                flag=1               
                cengshu=sheet2.cell(row2,2).value
                suoyou=sheet2.cell(row2,3).value
                time=sheet2.cell(row2,4).value
                zhiliang=sheet2.cell(row2,5).value
                jiegou=sheet2.cell(row2,6).value
                gn1=sheet2.cell(row2,7).value
                gn2=sheet2.cell(row2,8).value
                sheet.write(row1,19,cengshu)
                sheet.write(row1,20,suoyou)
                sheet.write(row1,21,time)
                sheet.write(row1,22,zhiliang)
                sheet.write(row1,23,jiegou)
                sheet.write(row1,24,gn1)
                sheet.write(row1,25,gn2)
                               
            else:
               hehe=''
               continue
        if flag==0:
           print("未匹配门牌号",menpaihao1)
           m=m+1
      
xls.save(r'C:\Users\33590\Desktop\0329hehe.xls')  
print("共计",m)