"数据库操作"
# import pymysql
# import xlrd
# import xlwt
# from xlutils.copy import copy
#
# database = pymysql.connect(host="localhost",user="root",
#                           password="123123",database='db',charset='utf8')  #连接数据库
# cursor = database.cursor()  #初始化指针
#
# 增
# sql = "INSERT INTO data (date,company,province,price,weight) VALUES ('2021-6-18','永济粮食','山西','1000','61.8')"
# cursor.execute(sql)
# database.commit()   #对存储的数据修改后，一定需要commit（）
#
# 改
# sql = "UPDATE data set date='2022-06-22' WHERE date='2021-06-22' and id='1775';"
# cursor.execute(sql)
# database.commit()
# database.close()
#
# 查
# sql = "SELECT company,sum(weight) FROM data WHERE date='2021-6-22' GROUP BY company;" #以公司为分组，查询质量总和
# cursor.execute(sql)
# result = cursor.fetchall()
# print(result)
#
# 删
# sql = "DELETE FROM data WHERE date='2021-06-22'"
# cursor.execute(sql)
# database.commit()
# database.close()
#
# 利用MySQL自动生成统计报表
# sql = "SELECT company ,COUNT(company),SUM(weight),SUM(weight*price) FROM data GROUP BY company"
# cursor.execute(sql)
# result = cursor.fetchall()
# print(result)
#
# for i in result:
#     if i[0] == '张三粮配':
#         a_num = i[1]
#         a_weight = i[2]
#         a_total_price = i[3]
#     elif i[0] == '李四粮食':
#         b_num = i[1]
#         b_weight = i[2]
#         b_total_price = i[3]
#     elif i[0] == '王五小麦':
#         c_num = i[1]
#         c_weight = i[2]
#         c_total_price = i[3]
#     elif i[0] == '赵六麦子专营':
#         d_num = i[1]
#         d_weight = i[2]
#         d_total_price = i[3]
#
# tem_excel = xlrd.open_workbook('7月下旬统计表.xls',formatting_info=True)
# tem_sheet = tem_excel.sheet_by_index(0)
#
# new_excel = copy(tem_excel)
# new_sheet = new_excel.get_sheet(0)
#
# style = xlwt.XFStyle()
#
# new_sheet.write(2,1,a_num)
# new_sheet.write(2,2,a_weight)
# new_sheet.write(2,3,a_total_price)
# new_sheet.write(3,1,b_num)
# new_sheet.write(3,2,b_weight)
# new_sheet.write(3,3,b_total_price)
# new_sheet.write(4,1,c_num)
# new_sheet.write(4,2,c_weight)
# new_sheet.write(4,3,c_total_price)
# new_sheet.write(5,1,d_num)
# new_sheet.write(5,2,d_weight)
# new_sheet.write(5,3,d_total_price)
#
# new_excel.save('7月下旬统计表1.xls')
