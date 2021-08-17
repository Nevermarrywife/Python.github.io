import xlsxwriter

wb = xlsxwriter.Workbook("datas.xlsx")  #创建工作簿
sheet = wb.add_worksheet("data_sheet")  #创建工作表

cell_format = wb.add_format({'bold':True})  #创建格式对象并格式化
cell_format1 = wb.add_format()  #创建格式对象
cell_format1.set_font_size(14)  #格式化，字号
cell_format1.set_bold() #加粗
cell_format1.set_align('center')    #居中
cell_format2 = wb.add_format()
cell_format2.set_bg_color('#FF00FF')

sheet.write(0, 0, '2020年度财务统计',cell_format) #写入单元格
sheet.merge_range(1, 0, 2, 2, "一季度销售统计",cell_format1)   #写入单元格

data = (
    ["一月份", 500, 450],
    ["二月份", 600, 650],
    ["三月份", 600, 550]
)

sheet.write_row(3, 0, ["月份", "预期销售额", "实际销售额"],cell_format2) #整行写入

for n, i in enumerate(data):    #遍历数据并且整行写入
    sheet.write_row(n + 4, 0, i)
sheet.write(7, 1, "=sum(B5:B7)")
sheet.write(7, 2, "=sum(C5:C7)")
sheet.write_url(9, 0, "http://www.baidu.com", string="更多资料")
sheet.insert_image(10, 0, "ship.bmp")

wb.close()  #关闭
