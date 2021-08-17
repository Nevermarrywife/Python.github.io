# import xlrd
#
# data = xlrd.open_workbook(r"D:\auto\test1.xlsx") #将test1.xlsx工作簿对象返回给data，默认加载所有工作表
# print(data.sheet_loaded(0))  #输出顺序为0的工作表的加载结果
# data.unload_sheet(0)    #卸载对工作表为0顺序的加载
# print(data.sheet_loaded(0)) #结果为False
# print(data.sheets()) #获取全部工作表，sheets()返回的是一个列表
# print(data.sheets()[0]) #通过列表获取工作表
# print(data.sheet_by_index(0))   #通过索引获取工作表
# print(data.sheet_by_name("Sheet1")) #通过工作表名称获取
# print(data.sheet_names())   #获取所有工作表的名字
# print(data.nsheets) #返回工作表的数量

#操作EXCEL行
# sheet1 = data.sheet_by_index(0) #获取第一个工作表对象返回给sheet1
# print(sheet1.nrows) #获取工作表的有效行数
# print(sheet1.row(1))    #获取第2行单元格对象组成的列表
# print(sheet1.row_types(1))  #获取第2行单元格的数据类型，1代表文本、2代表数字
# print(sheet1.row(1)[2].value)   #获取第2行第3列单元格的内容
# print(sheet1.row_values(1)) #获取指定行的内容列表
# print(sheet1.row_len(1))    #获取指定行的长度

#操作EXCEL列
# sheet1 = data.sheet_by_index(0)
# print(sheet1.ncols) #获取有效列数
# print(sheet1.col(1))    #获取第2列单元格对象的列表
# print(sheet1.col_types(1))  #获取第2列单元格的数据类型
# print(sheet1.col(1)[2].value)   #获取第2列第3行单元格的内容
# print(sheet1.col_values(1)) #获取第2列单元格的内容列表

#操作EXCEL单元格
# sheet1 = data.sheet_by_index(0)
# print(sheet1.cell(1,2)) #获取第2行第3列单元格对象
# print(sheet1.cell_type(1,2))    #获取单元格数据类型
# print(sheet1.cell(1,2).ctype)   #获取单元格数据类型
# print(sheet1.cell(1,2).value)   #获取单元格内容
# print(sheet1.cell_value(1,2))   #获取单元格内容