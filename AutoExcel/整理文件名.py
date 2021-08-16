import os
import xlwt

file_dir = 'd:/'   #要整理的目录
os.listdir(file_dir)    #生成文件名
new_workbook = xlwt.Workbook()
worksheet = new_workbook.add_sheet('new_test')  #创建工作簿和工作表

n = 0
for i in os.listdir(file_dir):  #遍历文件名
    worksheet.write(n,0,i)
    n +=1
new_workbook.save('file_name.xls')  #保存为这个工作表