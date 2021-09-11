
import os               # 操作Windows文件用
import time
import random
#import hashlib
import rsa
import xlrd
import xlwt
import sys

## 数据库
import sqlite3

## 图形化模块
from tkinter import *
from tkinter import filedialog  # 选择文件用
import tkinter.messagebox       # 弹出提示对话框
from tkinter import ttk         # 下拉菜单控件在ttk中

top = tkinter.Tk()
top.title('REPORT_DAY')
width=800
height=600
screenwidth = top.winfo_screenwidth()
screenheight = top.winfo_screenheight ()
alignstr ='%dx%d+%d+%d' %(width,height,(screenwidth-800)/2,(screenheight-600)/2)
top.geometry(alignstr)
top.resizable(width=False,height=False)
global labelframe_edk
#labelframe_edk = LabelFrame(top, text='当日通报',width=700,height=330)  #通报报表显示框
labelframe_button = LabelFrame(top, text='按钮框',width=100,height=330)  #操作按钮框
#labelframe_edk.grid(row=0,column=0,sticky='NW')
labelframe_button.grid(row=0,column=1,sticky='NW')

print("程序开始执行")
global kuandai_flag
kuandai_flag = IntVar()    #表示是否是插入宽带提取SQL
global jifen_flag
jifen_flag = IntVar()

def cmd_insert():
         global kuandai_flag
         global jifen_flag
         sql_name = tb_type.get()
         print(sql_name)
         query_cmd1="""select a.city_code,a.city_name,count(distinct phone_no)
                       from city_code_table a 
                       LEFT OUTER JOIN day_report b on a.city_name=b.city_name and substr(phone_no,1,1)='1'
                       and  b.op_no in (
                  """
         query_cmd2 = op_no.get()
         query_cmd3 = """) and b.op_bak like '%"""
         query_cmd4 = gj_word.get()
         query_cmd5 ="""%' group by a.city_code,a.city_name order by a.city_code"""
         query_cmd=query_cmd1+query_cmd2+query_cmd3+query_cmd4+query_cmd5   #无线指标统计SQL
        # print(query_cmd)
         query_cmd11="""select a.city_code,a.city_name,count(distinct 宽带账号)
                        from city_code_table a
                        LEFT OUTER JOIN day_kuandai b on substr(a.city_name,1,2)=substr(b.区县,1,2)
                        and b.工单状态 in ("""
         query_cmd31=""") and b.业务类型 in ("""
         query_cmd51=""") group by a.city_code,a.city_name order by a.city_code """
         query_cmd21=query_cmd11+query_cmd2+query_cmd31+query_cmd4+query_cmd51 #宽带指标统计SQL
         kuandai_flag1 = kuandai_flag.get()
         jifen_flag1 = jifen_flag.get()
         print('宽带标志:',kuandai_flag1)
         print('积分标志：',jifen_flag1)
         if (sql_name=='' or query_cmd2=='' or query_cmd4==''):
                  tkinter.messagebox.showinfo(message="指标名称、关键字及业务代码均不能为空" )
         else:
                  if kuandai_flag1 ==0 and jifen_flag1==0:
                           sql_cmd_code =f'INSERT INTO data_querysql_talbe(querysql_name,sql_cmd) VALUES ("{sql_name}", "{query_cmd}");'
                           conn = sqlite3.connect('tb_report_db.DB')
                           #print('无线发展指标:',kuandai_flag1)
                           cursor = conn.cursor()
                           cursor.execute(sql_cmd_code)
                           conn.commit()
                           cursor.close()
                           conn.close()
                           print('无线发展指标提取SQL添加成功')
                  if kuandai_flag1 ==1 and jifen_flag1==0:
                           sql_cmd_code =f'INSERT INTO data_querysql_talbe(querysql_name,sql_cmd) VALUES ("{sql_name}", "{query_cmd21}");'
                           conn = sqlite3.connect('tb_report_db.DB')
                           #print('宽带发展指标：',kuandai_flag1)
                           cursor = conn.cursor()
                           cursor.execute(sql_cmd_code)
                           conn.commit()
                           cursor.close()
                           conn.close()
                           print('宽带发展指标提取SQL添加成功')
                  if kuandai_flag1 ==0 and jifen_flag1==1:
                           sql_cmd_code =f'INSERT INTO activity_jifen(activity_name,jifen,jifen_value) VALUES ("{sql_name}", "{query_cmd2}","{query_cmd4}");'
                           conn = sqlite3.connect('tb_report_db.DB')
                           print('积分指标：',jifen_flag1)
                           cursor = conn.cursor()
                           cursor.execute(sql_cmd_code)
                           conn.commit()
                           cursor.close()
                           conn.close()
                           print('积分活动添加成功')
                  if kuandai_flag1 ==1 and jifen_flag1==1:
                       tkinter.messagebox.showinfo(message='宽带标志和积分标志不能同时打勾')
    

def cmd_query_item():
         text_querysql.delete(0.0, END)
         text_querysql.focus_set()
         sql_cmd ="""select * from  data_querysql_talbe"""
         conn = sqlite3.connect('tb_report_db.DB')
         print('查询指标SQL语句成功')
         cursor = conn.cursor()
         jieguo = cursor.execute(sql_cmd).fetchall()
         for i in range(len(jieguo)):
                  text_querysql.insert(END,jieguo[i])
                  text_querysql.insert(END,'\n')

def delete_text():
         text_querysql.delete(0.0, END)
         text_querysql.focus_set()
         tb_type.delete(0,END)
         tb_type.focus_set()
         op_no.delete(0,END)
         op_no.focus_set()
         gj_word.delete(0,END)
         gj_word.focus_set()
def delete_item():
         sql_name = tb_type.get()
         if sql_name == '':
                   tkinter.messagebox.showinfo(message='请输入指标名称')
         else:
                  sql_cmd =f'DELETE FROM data_querysql_talbe where querysql_name= "{sql_name}";'
                  conn = sqlite3.connect('tb_report_db.DB')
         #print('连接数据库成功')
                  cursor = conn.cursor()
                  cursor.execute(sql_cmd)
                  conn.commit()
                  cursor.close()
                  conn.close()
                  tkinter.messagebox.showinfo(message="删除指标：%s完成" %tb_type.get() )




sqlcmd_edk = tkinter.LabelFrame(top,width = 700, text='指标编辑区')


#指标增加、修改模块

item_edk = tkinter.LabelFrame(sqlcmd_edk,width=700)
item_edk.grid(row =0,column=0,sticky =S+W+E+N)

tkinter.Label(item_edk,text='指标名称：',width=12).grid(row =0,column = 0,sticky='NW')

#定义文本输入框
text_querysql = tkinter.Text(sqlcmd_edk, height =10)
#text_querysql.insert(END,'先清空文本框')
scrollbar_text = tkinter.Scrollbar(sqlcmd_edk,orient="vertical",command=text_querysql.yview)
text_querysql.config(yscrollcommand = scrollbar_text.set)
text_querysql.grid(row = 1,column =0,sticky=S+W+E+N)
scrollbar_text.grid(row = 1,column =1,sticky=S+W+E+N )
sqlcmd_edk.grid(row=2,column=0,sticky=S+W+E+N)

tb_type = tkinter.Entry(item_edk,width=30)   #输入新增指标名称
tb_type.grid(row =0,column = 1,sticky=W)

tkinter.Label(item_edk,text='业务代码：',width=12).grid(row =1,column = 0,sticky=W)

op_no = tkinter.Entry(item_edk,width=30)   #输入业务代码
op_no.grid(row =1,column = 1,sticky='NW')

tkinter.Label(item_edk,text='关键字：',width=12).grid(row =2,column = 0,sticky=W)

gj_word = tkinter.Entry(item_edk,width=30)   #输入关键字
gj_word.grid(row =2,column = 1,sticky=W)

Button(item_edk, text=('  '+'添加指标'+'  '), command=cmd_insert).grid(row=1, column=2,sticky=W)
tkinter.Label(item_edk,text='宽带发展指标，请打勾，业务代码-工单状态',fg='red').grid(row =0,column = 3,sticky=W)
tkinter.Checkbutton(item_edk, text=('宽带标志'), variable=kuandai_flag).grid(row=0, column=2,sticky='NE')
Button(item_edk, text=('  '+'查询指标'+'  '), command=cmd_query_item).grid(row=2, column=2,sticky = W)
Button(item_edk, text=('  '+'清空文本'+'  '), command=delete_text).grid(row=1, column=3,sticky = W)
tkinter.Checkbutton(item_edk, text=('积分标志'), variable=jifen_flag).grid(row=1, column=3,sticky='NE')
Button(item_edk, text=('  '+'删除指标'+'  '), command=delete_item).grid(row=2, column=3,sticky=W)


global export_field_title
global export_data_list
global bar_flag
global field,canvas1

field_zb = {}  #KEY=框架坐标 value=初始值
data_zb = {}
data_zb_row = IntVar()
data_zb_column = IntVar()
#canvas1 = tkinter.Canvas (labelframe_edk,width=700,height=330,bg='#00CED1')

def frame_clear(frame_name):                  #清除所有框架
         for x in frame_name.winfo_children():
                  print(x)
                  x.destroy()
def show_data(field_title,data_list):
         data_row = 14        #数据为十四行，为固定值
         labelframe_edk = LabelFrame(top, text='当日通报',width=700,height=330)  #通报报表显示框
         labelframe_edk.grid(row=0,column=0,sticky='NW')
         canvas1 = tkinter.Canvas (labelframe_edk,width=700,height=330,bg='#00CED1')
         frame_in_canvas = Frame(canvas1)
         frame_in_canvas.grid(row=0,column=0)
         field = Frame(frame_in_canvas)
         field.grid(row = 0,column = 0,sticky = 'NW')
         data = Frame(frame_in_canvas)
         data.grid(row =1,column = 0,sticky = 'NW')
         def update_scrollregion(event):
             canvas1.configure(scrollregion=canvas1.bbox("all"))
         scrollbar_w = Scrollbar(labelframe_edk,orient=HORIZONTAL)
         scrollbar_w.grid( row = 2,column = 0, sticky ="ew" )
         scrollbar_w.config(command=canvas1.xview)
         canvas1.config(xscrollcommand=scrollbar_w.set,scrollregion=canvas1.bbox("all"))
         canvas1.create_window(0,0,window =frame_in_canvas,anchor=NW)
         canvas1.grid(row = 0,column =0,sticky = "ew")
         canvas1.create_window(0,0,window =frame_in_canvas,anchor=NW)
         canvas1.bind("<Configure>",update_scrollregion) #绑定事件




         #清除全局字典中的内容
         field_zb.clear()

         #第一个字段
         field_index=0
         print ("字段序列长度",len(field_title))
         print("这个field_title:",field_title)
         for field_name in field_title:
                  print(field_name)
                  init_value = str(field_title[field_index])
                  #print(init_value)
                  field_zb[(0,field_index)]=init_value
                  input_entry = Entry(field,width = 20,bg='#00BFFF')
                  input_entry.insert(0,init_value)
                  input_entry.grid(row = 0,column = field_index,sticky='NW')
                  field_index+=1
                  print(field_index)

         #提取报表数据
         print("当前字段列数",str(field_index))
      
         for d_row in range(0,data_row):           #将区县写入展示框
                           init_value = str(data_list[0][d_row][1])
                           data_zb[(d_row,0)] = init_value
                           data_entry = Entry(data,width = 20)
                           data_entry.insert(0,init_value)
                           data_entry.grid(row =d_row+1,column = 0,sticky = 'W')

         for data_column in range(0,field_index-1):
                  for d_row in range(0,data_row):
                           init_value = str(data_list[data_column][d_row][2])
                           data_zb[(d_row,0)] = init_value
                           data_entry = Entry(data,width = 20)
                           data_entry.insert(0,init_value)
                           data_entry.grid(row =d_row+1,column = data_column+1,sticky = 'W')
#定义导出数据的全局变量
def open_database():
         global export_field_title
         global export_data_list
         global canvas1
         export_data_list=[]
         export_field_title=[]
         sql_cmd = """select * from data_querysql_talbe""" #提取指标名称
         conn = sqlite3.connect('tb_report_db.DB')
         print('进行指标展示')
         cursor = conn.cursor()
         field_list = cursor.execute(sql_cmd).fetchall()
         field_dict = {}
         for i in field_list:              #将查询结果转化为字典{指标名称：SQL语句}
                  field_dict[i[0]]=i[1]
         #将数据库的字段转化为列表，进行前台展示
         field_title = ['区县']
         sql_list = []
         for i_item in field_dict:
                  field_title.append(i_item)
                  sql_list.append(field_dict[i_item])
         print("字段列表",field_title)

         data_list=[]     #提取数据结果
         for i_sqlcmd in sql_list:
                  qu = cursor.execute(i_sqlcmd).fetchall()
                  data_list.append(qu)
         print("数据列表",data_list)
         export_field_title = field_title
         print('export_field_title:',export_field_title)
         export_data_list = data_list
         show_data(export_field_title,export_data_list)  #调用函数show_data()提取通报数据
 
         #print('export_data_list:',export_data_list)
         cursor.close
         conn.close
         #return export_field_title,export_data_list
def zhibiao_list(zb_list=[]):
         sql_cmd="""select querysql_name from data_querysql_talbe order by querysql_name"""
         conn = sqlite3.connect('tb_report_db.DB')
         print('连接数据库')
         cursor = conn.cursor()
         data_list = cursor.execute(sql_cmd).fetchall()
         for  i in data_list:
                  zb_list.append(i[0])
         cursor.close
         conn.close
         return zb_list
def input_table_kuandai(new_file_name):
         workbook = xlrd.open_workbook(new_file_name) #打开excel文件
         sheet_content = workbook.sheet_by_index(0) #取sheet1
         row_num = sheet_content.nrows  #取文件行数
         db = sqlite3.connect('tb_report_db.DB')  #打开数据库IU
         cursor = db.cursor()
         data_list = []  #定义暂存列表来存放数据
         table_name = 'day_kuandai'
         num =0
         for row_i in range(1,row_num):
                  row_data = sheet_content.row_values(row_i)
                  row_value = (row_data[0],row_data[1],row_data[2],row_data[3],row_data[4],
                               row_data[5],row_data[6],row_data[7],row_data[8],
                               row_data[9],row_data[10],row_data[11],row_data[12],
                               row_data[13],row_data[14],row_data[15],row_data[16],
                               row_data[17],row_data[18],row_data[19],row_data[20],
                               row_data[21],row_data[22],row_data[23])
                  data_list.append(row_value)
                  num+=1
         sql="""insert into""" +""" """+table_name  +"""(工单编号,客户名称,工单类型,L单号,
                              业务类型,宽带账号,IPTV账号,IMS账号,带宽,地市,区县,小区名称,
                              装机地址,处理人,工单状态,预改约次数,受理营业员名字,建单时间,首响时间,
                              预约时间,缓装时间,改约时间,回复时间,已用时间 )
                       VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"""
         cursor.executemany(sql,data_list)
         data_list.clear()
         tkinter.messagebox.showinfo(message="宽带报表导入了%d行数据" %num)
         db.commit()
         cursor.close()
         db.close()
def input_table(new_file_name):
         workbook = xlrd.open_workbook(new_file_name) #打开excel文件
         sheet_content = workbook.sheet_by_index(0) #取sheet1
         row_num = sheet_content.nrows  #取文件行数
         db = sqlite3.connect('tb_report_db.DB')  #打开数据库IU
         cursor = db.cursor()
         data_list = []  #定义暂存列表来存放数据
         table_name = 'day_report'
         num =0
         for row_i in range(3,row_num):
                  row_data = sheet_content.row_values(row_i)
                  if row_data[7]!='':
                           row_value = (row_data[1],row_data[2],row_data[3],row_data[4],
                                        row_data[5],row_data[6],row_data[7],row_data[8],
                                        row_data[9],row_data[10],row_data[11],row_data[12],
                                        row_data[13]
                                        )
                           data_list.append(row_value)
                           num+=1

         sql="""insert into""" +""" """+table_name  +"""(city_name,chnl_name,
                              login_no,login_name ,op_no,op_name,phone_no,op_time,
                              op_flow,chnl_type,op_bak ,ifwd,op_count )
                       VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)"""
         cursor.executemany(sql,data_list)
         data_list.clear()
         tkinter.messagebox.showinfo(message="无线报表导入了%d行数据" %num)
         db.commit()
         cursor.close()
         db.close()
def input_table_xiezhuan(new_file_name):
         workbook = xlrd.open_workbook(new_file_name) #打开excel文件
         sheet_content = workbook.sheet_by_index(0) #取sheet1
         row_num = sheet_content.nrows  #取文件行数
         db = sqlite3.connect('tb_report_db.DB')  #打开数据库IU
         cursor = db.cursor()
         data_list = []  #定义暂存列表来存放数据
         table_name = 'day_xiezhuan'
         num =0
         for row_i in range(3,row_num):
                  row_data = sheet_content.row_values(row_i)
                  if row_data[1]!='':
                           row_value = (row_data[1],row_data[2],row_data[3],row_data[4],
                                        row_data[5],row_data[6],row_data[7],row_data[8],
                                        row_data[9])
                           data_list.append(row_value)
                           num+=1
         sql="""insert into""" +""" """+table_name  +"""(phone_no,类型,操作时间,操作工号,操作流水,号码归属地市,city_name,号码归属网点,数量)VALUES(?,?,?,?,?,?,?,?,?)"""
         cursor.executemany(sql,data_list)
         data_list.clear()
         tkinter.messagebox.showinfo(message="携转报表导入了%d行数据" %num)
         db.commit()
         cursor.close()
         db.close()
def input_table_znyj(new_file_name):
         workbook = xlrd.open_workbook(new_file_name) #打开excel文件
         sheet_content = workbook.sheet_by_index(0) #取sheet1
         row_num = sheet_content.nrows  #取文件行数
         db = sqlite3.connect('tb_report_db.DB')  #打开数据库IU
         cursor = db.cursor()
         data_list = []  #定义暂存列表来存放数据
         table_name = 'day_znyj'
         num =0
         for row_i in range(3,row_num):
                  row_data = sheet_content.row_values(row_i)
                  if row_data[1]!='':
                           row_value = (row_data[0],row_data[1],row_data[2],row_data[3],row_data[4]
                                        ,row_data[5],row_data[6],row_data[7],row_data[8],row_data[9],row_data[10]
                                        ,row_data[11],row_data[12],row_data[13],row_data[14],row_data[15],row_data[16]
                                        ,row_data[17],row_data[18],row_data[19],row_data[20],row_data[21],row_data[22]
                                        ,row_data[23],row_data[24],row_data[25])
                           data_list.append(row_value)
                           num+=1
         sql="""insert into""" +""" """+table_name  +"""
(单位名称,操作工号,销售营业厅代码,活动名称,设备类别,品牌名称,终端类型,终端大类,预约方式,是否裸机销售,终端名称,合作类型,机型代码,机器串号,
手机号码,销售工号,销售工号归属营业厅,IPAD销售工号,导购方式,数量,成本价,现金购机款,和包购机款,终端成本溢价,终端成本实时溢价,信用贷款)
VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"""
         cursor.executemany(sql,data_list)
         data_list.clear()
         tkinter.messagebox.showinfo(message="携转报表导入了%d行数据" %num)
         db.commit()
         cursor.close()
         db.close()
         
global inputfile_flag
inputfile_flag =[]
new_file_name = StringVar()
def input_data():    #导入通报报表
         global inputfile_flag
         file_flag=False
         new_file_name = filedialog.askopenfilename()
         labelframe_button['text'] = "已选择:"+new_file_name
         file_flag=tkinter.messagebox.askyesno(title='',message='请核对无线文件是否选择正确')
         if file_flag ==True:
                  inputfile_flag.append('无线')
                  input_table(new_file_name)
                  labelframe_button['text'] = "无线操作数据已导入"
         else:
                  tkinter.messagebox.showinfo(message="请重新选择正确文件" )
def input_data_kuandai():    #导入宽带报表
         global inputfile_flag
         file_flag=False
         new_file_name = filedialog.askopenfilename()
         labelframe_button['text'] = "已选择:"+new_file_name
         file_flag=tkinter.messagebox.askyesno(title='',message='请核对宽带文件是否选择正确')
         if file_flag ==True:
                  inputfile_flag.append('宽带')
                  input_table_kuandai (new_file_name)
                  labelframe_button['text'] = "宽带数据已导入"
                  if len(inputfile_flag)==2:
                           tkinter.messagebox.showinfo(message="无线和宽带数据已全部导入" )
                           labelframe_button['text'] = "无线和宽带数据已全部导入"
         else:
                  tkinter.messagebox.showinfo(message="请重新选择正确文件" )
def input_data_xiezhuan():    #导入携转报表
         global inputfile_flag
         file_flag=False
         new_file_name = filedialog.askopenfilename()
         labelframe_button['text'] = "已选择:"+new_file_name
         file_flag=tkinter.messagebox.askyesno(title='',message='请核对携转文件是否选择正确')
         if file_flag ==True:
                  inputfile_flag.append('携转')
                  input_table_xiezhuan (new_file_name)
                  labelframe_button['text'] = "携转数据已导入"
                  if len(inputfile_flag)==2:
                           tkinter.messagebox.showinfo(message="所有数据已全部导入" )
                           labelframe_button['text'] = "所有数据已全部导入"
         else:
                  tkinter.messagebox.showinfo(message="请重新选择正确文件" )
def input_data_znyj():    #导入携转报表
         global inputfile_flag
         file_flag=False
         new_file_name = filedialog.askopenfilename()
         labelframe_button['text'] = "已选择:"+new_file_name
         file_flag=tkinter.messagebox.askyesno(title='',message='请核对智能文件是否选择正确')
         if file_flag ==True:
                  inputfile_flag.append('智能硬件')
                  input_table_znyj (new_file_name)
                  labelframe_button['text'] = "智能硬件数据已导入"
                  if len(inputfile_flag)==2:
                           tkinter.messagebox.showinfo(message="所有数据已全部导入" )
                           labelframe_button['text'] = "所有数据已全部导入"
         else:
                  tkinter.messagebox.showinfo(message="请重新选择正确文件" )


def export_data():         #导出到excel
         book = xlwt.Workbook()
         i=0
         sheet = book.add_sheet('导出数据')
         #export_field_title,export_data_list = open_database()
         #export_field_title=export_field_title
         print("export_data_list",export_data_list)
         list_flag = len(export_data_list)   #有几个指标

         print(export_field_title)
         print(export_data_list)
         print("row",len(export_data_list[0]))
         print("col",len(export_data_list[0][0]))

         for header in export_field_title:
                  print("header:",header)
                  sheet.write(0,i,header)
                  i+=1


         zb_flag=1    #指标开始
         for zb in export_data_list:
                  if zb_flag ==1:
                           for row in range(1,len(zb)+1):
                                    for col in range(1,len(zb[row-1])):
                                             print(zb[row-1][col],end = '')
                                             sheet.write(row,col-1,zb[row-1][col])
                           #zb_flag+=1
                  else:
                           for row in range(1,len(zb)+1):
                                    print("第二个指标")
                                    print(zb[row-1][2])
                                    sheet.write(row,zb_flag,zb[row-1][2])
                  zb_flag+=1
         www = time.localtime()
         file_name_time ="\\"+ str(www[0])+str(www[1])+str(www[2])+str(www[3])+str(www[4])+str(www[5])+'日报数据.xls'
         file_path =os.path.abspath(file_name_time)
         print(file_path)
         book.save(file_path)
         tkinter.messagebox.showinfo(message="报表成功导入：%s完成" %file_path )

def delete_data():
         conn = sqlite3.connect('tb_report_db.DB')
         cursor = conn.cursor()
         sql_cmd = """delete  from day_report"""
         sql_cmd2= """delete from day_kuandai"""
         sql_cmd3= """delete from day_xiezhuan"""
         sql_cmd4 = """delete from day_znyj"""
         
         cursor.execute(sql_cmd)
         cursor.execute(sql_cmd2)
         cursor.execute(sql_cmd3)
         cursor.execute(sql_cmd4)
         conn.commit()
         cursor.close()
         conn.close()
         tkinter.messagebox.showinfo(message="历史数据已清空")

def tichu_data():
         conn = sqlite3.connect('tb_report_db.DB')
         cursor = conn.cursor()
         sql_cmd = """delete  from day_report where (op_bak like '%退订%' or op_bak like'%客户进行家庭融合群开户%')"""
         cursor.execute(sql_cmd)
         conn.commit()
         cursor.close()
         conn.close()
         tkinter.messagebox.showinfo(message="退订和家庭融合群开户数据已删除")


Button(labelframe_button, text=('  '+'清空历史数据'+'  '), command=delete_data).grid(row=1, column=1)
Button(labelframe_button, text=('  '+'导入无线数据'+'  '), command=input_data).grid(row=2, column=1)
Button(labelframe_button, text=('  '+'导入宽带数据'+'  '), command=input_data_kuandai).grid(row=3, column=1)
Button(labelframe_button, text=('  '+'导入携转数据'+'  '), command=input_data_xiezhuan).grid(row=4, column=1)
Button(labelframe_button, text=('  '+'导入智能数据'+'  '), command=input_data_znyj).grid(row=5, column=1)
Button(labelframe_button, text=('  '+'剔除非法数据'+'  '), command=tichu_data).grid(row=6, column=1)
Button(labelframe_button, text=('  '+'提取通报报表'+'  '), command=open_database).grid(row=7, column=1)
Button(labelframe_button, text=('  '+'导出通报报表'+'  '), command=export_data).grid(row=8, column=1)


def create_database():
         conn = sqlite3.connect('tb_report_db.DB')
         print('连接数据库')
         cursor = conn.cursor()
         create_day_report_sql="""
                               CREATE TABLE IF NOT EXISTS day_report
                              (city_name varchar(20),chnl_name varchar,
                              login_no varchar(7),login_name varchar,
                              op_no varchar(4),op_name varchar,
                              phone_no varchar(15),op_time var_char,op_flow varchar(15),
                              chnl_type var_char,op_bak varchar,
                              ifwd varchar(6),op_count varchar(2))
                              """
         cursor.execute(create_day_report_sql)
         create_day_kuandai_sql="""
                               CREATE TABLE IF NOT EXISTS day_kuandai
                              (工单编号 varchar(20),客户名称 varchar,
                              工单类型 varchar(8),L单号 varchar(20),
                              业务类型 varchar(30),宽带账号 varchar(15),
                              IPTV账号 varchar(15),IMS账号 var_char(15),带宽 varchar(10),
                              地市 var_char(8),区县 varchar(10),
                              小区名称 varchar,装机地址 varchar,处理人 varchar,工单状态 varchar(10),
                              预改约次数 integer,受理营业员名字 varchar,建单时间 varchar,首响时间 varchar,
                              预约时间 varchar,缓装时间 varchar,改约时间 varchar,回复时间 varchar,已用时间 varchar)
                              """
         cursor.execute(create_day_kuandai_sql)
         create_day_znyj_sql="""
                               CREATE TABLE IF NOT EXISTS day_znyj
                              (单位名称 varchar,操作工号 varchar,
                              销售营业厅代码 varchar,活动名称 varchar,
                              设备类别 varchar,品牌名称 varchar,
                              终端类型 varchar,终端大类 var_char,
                              预约方式 varchar,是否裸机销售 var_char,
                              终端名称 varchar,合作类型 varchar,
                              机型代码 varchar,机器串号 varchar,
                              手机号码 varchar,销售工号 vachar,
                              销售工号归属营业厅 varchar,IPAD销售工号 varchar,
                              导购方式 varchar,数量 integer,
                              成本价 float,现金购机款 float,
                              和包购机款 float,终端成本溢价 float,
                              终端成本实时溢价 float,信用贷款 float)
                              """
         cursor.execute(create_day_znyj_sql)
         create_jifen_hd_sql = """CREATE TABLE IF NOT EXISTS activity_jifen(activity_name varchar,jifen integer,jifen_value float)
                                          """
         cursor.execute(create_jifen_hd_sql)

         create_xz_sql="""CREATE TABLE IF NOT EXISTS day_xiezhuan
                                    (phone_no varchar,类型 varchar,操作时间 varchar,
                                    操作工号 varchar,操作流水 varchar,号码归属地市 varchar,
                                    city_name varchar(20),号码归属网点 varchar,数量 varchar(2))
                               """
         cursor.execute(create_xz_sql)

         sql2 = """CREATE TABLE IF NOT EXISTS city_code_table
                  (city_name varchar(20),city_code integer,login_gs varchar)
         """
         cursor.execute(sql2)

         sql3 = """CREATE TABLE IF NOT EXISTS data_querysql_talbe
                  (querysql_name varchar,sql_cmd varchar)
         """
         print(sql3)
         cursor.execute(sql3)
         sql_citycodesql = "select * from city_code_table"
         city = cursor.execute(sql_citycodesql).fetchall()
         if len(city)<5:
                  city_list=[('盐湖营业部',1,'ja'),('芮城分公司',2,'jb'),('平陆分公司',3,'jc'),('临猗分公司',4,'jd'),('万荣分公司',5,'je'),('河津分公司',6,'jf'),
                    ('稷山分公司',7,'jg'),('垣曲分公司',8,'jh'),('绛县分公司',9,'ji'),('闻喜分公司',10,'jj'),('新绛分公司',11,'jk'),('永济分公司',12,'jl'),
                    ('夏县分公司',13,'jm'),('运城网上商城营业部',14,'jw')]
                  sql_insertcitycode = "insert into city_code_table(city_name,city_code,login_gs) values(?,?,?)"
                  cursor.executemany(sql_insertcitycode,city_list)
                  city_list.clear()
         else:
                  print("数据完备")

         cmd_sql = "select * from sqlite_master"
         list_table = cursor.execute(cmd_sql).fetchall()
         print(list_table)

        #插入积分统计语句
         jifen_sql = """ select a.city_code,a.city_name,sum(c.jifen_value)
                       from city_code_table a
                       LEFT OUTER JOIN day_report b on a.city_name=b.city_name and substr(phone_no,1,1)='1'
                       and  b.op_no in ('1147','4035')
                       LEFT OUTER JOIN activity_jifen c on b.op_bak=c.activity_name 
                       group by a.city_code,a.city_name order by a.city_code"""
         sql_jifen_code =f'INSERT INTO data_querysql_talbe(querysql_name,sql_cmd) VALUES ("积分价值", "{jifen_sql}");'
         cursor.execute(sql_jifen_code)
         #插入携入语句：
         xieru_sql = """ select a.city_code,a.city_name,count(distinct phone_no)
                       from city_code_table a
                       LEFT OUTER JOIN day_xiezhuan b on a.city_name=b.city_name and substr(phone_no,1,1)='1'
                       and 类型='已携入'
                       group by a.city_code,a.city_name order by a.city_code"""
         sql_xieru_code =f'INSERT INTO data_querysql_talbe(querysql_name,sql_cmd) VALUES ("携入", "{xieru_sql}");'
         cursor.execute(sql_xieru_code)
         #插入携出语句：
         xiechu_sql = """ select a.city_code,a.city_name,count(distinct phone_no)
                       from city_code_table a
                       LEFT OUTER JOIN day_xiezhuan b on a.city_name=b.city_name and substr(phone_no,1,1)='1'
                       and 类型='已携出'
                       group by a.city_code,a.city_name order by a.city_code"""
         sql_xiechu_code =f'INSERT INTO data_querysql_talbe(querysql_name,sql_cmd) VALUES ("携出", "{xiechu_sql}");'
         cursor.execute(sql_xiechu_code)

          #插入智能家居业务语句：
         zjywtxlj_sql = """ select a.city_code,a.city_name,sum(数量)
                       from city_code_table a
                       LEFT OUTER JOIN day_znyj b on a.login_gs=substr(b.操作工号,1,2)
                       and 终端类型='通信连接' 
                       group by a.city_code,a.city_name order by a.city_code"""
         sql_zjywtxlj_code =f'INSERT INTO data_querysql_talbe(querysql_name,sql_cmd) VALUES ("通信连接", "{zjywtxlj_sql}");'
         cursor.execute(sql_zjywtxlj_code)
          #插入智能组网语句：
         znzwafjk_sql = """ select a.city_code,a.city_name,sum(数量)
                       from city_code_table a
                       LEFT OUTER JOIN day_znyj b on a.login_gs=substr(b.操作工号,1,2)
                       and 终端类型 = '安防监控'
                       group by a.city_code,a.city_name order by a.city_code"""
         sql_znzwafjk_code =f'INSERT INTO data_querysql_talbe(querysql_name,sql_cmd) VALUES ("安防监控", "{znzwafjk_sql}");'
         cursor.execute(sql_znzwafjk_code)
         #插入智能组网语句：
         znzwlyq_sql = """ select a.city_code,a.city_name,sum(数量)
                       from city_code_table a
                       LEFT OUTER JOIN day_znyj b on a.login_gs=substr(b.操作工号,1,2)
                       and 终端类型 = '路由器网关'
                       group by a.city_code,a.city_name order by a.city_code"""
         sql_znzwlyq_code =f'INSERT INTO data_querysql_talbe(querysql_name,sql_cmd) VALUES ("路由器网关", "{znzwlyq_sql}");'
         cursor.execute(sql_znzwlyq_code)
          #插入智能组网语句：
         znzwyljy_sql = """ select a.city_code,a.city_name,sum(数量)
                       from city_code_table a
                       LEFT OUTER JOIN day_znyj b on a.login_gs=substr(b.操作工号,1,2)
                       and 终端类型 ='教育娱乐'
                       group by a.city_code,a.city_name order by a.city_code"""
         sql_znzwyljy_code =f'INSERT INTO data_querysql_talbe(querysql_name,sql_cmd) VALUES ("教育娱乐", "{znzwyljy_sql}");'
         cursor.execute(sql_znzwyljy_code)
         conn.commit()
         cursor.close()
         conn.close()
         tkinter.messagebox.showinfo(message="数据库创建完毕")
Button(labelframe_button, text=('  '+'创建新数据库'+'  '), command=create_database).grid(row=9, column=1)
top.mainloop()
