import xlwt

"初始化excel"
wb = xlwt.Workbook()    #创建工作簿

"设置参数及数据"
ws = wb.add_sheet("cny")    #创建工作表，并命名为cny
ws.write_merge(0,1,0,5,"2020年货币兑换表")    #合并前2行及前6列，并写入内容
data = (("Date","英镑","人民币","港币","日元","美元"),("01/01/2019",8.72251,1,0.877885,0.0062722,6.8759),
        ("02/01/2019",8.634922,1,0.876731,0.062773,6.8601))

"遍历data数据，并写入excel"
for i,item in enumerate(data):  #enumerate()将data序列组合成带索引的序列，并将索引赋给i
        for j,val in enumerate(item):
                ws.write(i+2,j,val)     #由于前2行已经有数据填充，因此i+2
                #print(f"i = {i},item = {item},val = {val}")

"在另一个工作表里写入图片"
wsimage = wb.add_sheet("images")        #创建新工作表

wsimage.insert_bitmap("ship.bmp",0,0)   #insert_bitmap函数插入文件，需要是bmp文件

"保存工作簿"
wb.save("2020-CNY.xls") #保存工作簿，并写入命名
