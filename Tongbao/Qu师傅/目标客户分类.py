
import os               # 操作Windows文件用
import xlrd
import xlwt
import re
import pandas as pd
## 数据库

## 图形化模块
from tkinter import *
from tkinter import filedialog  # 选择文件用
import tkinter.messagebox       # 弹出提示对话框
from tkinter import ttk         # 下拉菜单控件在ttk中

top = tkinter.Tk()
top.title('Excel文件分类')
width=800
height=600
screenwidth = top.winfo_screenwidth()
screenheight = top.winfo_screenheight ()
alignstr ='%dx%d+%d+%d' %(width,height,(screenwidth-800)/2,(screenheight-600)/2)
top.geometry(alignstr)
top.resizable(width=False,height=False)
print("程序开始执行")



def input_data():            #选择分类文件
    new_file = filedialog.askopenfilename()
    workbook = xlrd.open_workbook(new_file) #打开excel文件
    sheet_name = workbook.sheet_names() #获取sheet名称
    return new_file,sheet_name
         
def export_data():         #对文件进行分类并导出
   # new_file,sheet_name = input_data()
    new_file = filedialog.askopenfilename()
    dire = os.path.split(new_file)[0] #获取文件目录
    fg_file  = os.path.split(new_file)[1]
    fg_name=re.findall(r'(.+?)\.',fg_file)
    data = pd.read_excel(new_file)
    area_list =list(set(data['区县代码']))
    for i in area_list:
        writer = pd.ExcelWriter(dire+"/"+str(i)+'-'+fg_name[0]+".xlsx")
        df = data[data['区县代码'] == i]
        label1 = tkinter.Label(top,text='执行进度：'+str(i),width=30).grid(row =3,column = 1,sticky='NW')
        df.to_excel(writer,sheet_name=fg_name[0],index=False)
        writer.save()
        writer.close()
button2 = tkinter.Button(top, text=('  '+'开始分类'+'  '), command=export_data).grid(row=2, column=1)
label2 = tkinter.Label(top,text='使用说明：区县代码列的标题名称必须修改为“区县代码”；生成文件与分类文件在同一目录',width=100).grid(row =4,column = 1,sticky='NW')
top.mainloop()
