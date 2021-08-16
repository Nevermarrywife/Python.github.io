import xlrd
import xlwt
from xlutils.copy import copy

xls = xlrd.open_workbook('5.xls')
table = xls.sheet_by_index(0)

all_data = []
s = 0

for n in range(3, table.nrows):
    name = table.cell(n, 4).value
    work = table.cell(n, 11).value
    num = table.cell(n, 5).value
    type = table.cell(n, 6).value
    obj = table.cell(n,7).value
    data = {'name': name, 'work': work, 'num': num, 'type': type,'obj':obj}
    all_data.append(data)

a_point = 0
a_work = []
a_work_num = []
a_work_turn = []
b_point = 0
b_work = []
b_work_num = []
b_work_turn = []
c_point = 0
c_work = []
c_work_num = []
c_work_turn = []
d_point = 0
d_work = []
d_work_num = []
d_work_turn = []
e_point = 0
e_work = []
e_work_num = []
e_work_turn = []
f_point = 0
f_work = []
f_work_num = []
f_work_turn = []
g_point = 0
g_work = []
g_work_num = []
g_work_turn = []
h_point = 0
h_work = []
h_work_num = []
h_work_turn = []

k = 'SC2171-二季度家庭融合享优惠（副卡）'
g = 'SC2172-二季度家庭融合享优惠（家庭网）'
s = 'SC2187-运城移动客户感恩有礼回馈活动'

active = ['SC2171-二季度家庭融合享优惠（副卡）', 'SC2172-二季度家庭融合享优惠（家庭网）', 'SC2187-运城移动客户感恩有礼回馈活动']

points = [{'type': '2019积分兑换活动(3G全国流量)', 'cent': 1000},
          {'type': '2019积分兑换活动(9G全国流量)', 'cent': 3000},
          {'type': '2019年积分兑换直充包（流量特惠包1GB(24小时)）', 'cent': 5000}]

for i in all_data:
    if i['name'] == '张敏':
        c = 5
        a_work.append(i['work'])
        if i['num'] == '1000':
            a_work_num.append(i['work'])
        if i['type'] == '携号转网开户':
            a_work_turn.append(i['work'])
for m in a_work:
    for n in points:
        if m == n['type']:
            a_point += n['cent']

for i in all_data:
    if i['name'] == '贾洋洋':
        b_work.append(i['work'])
        if i['num'] == '1000':
            b_work_num.append(i['work'])
        if i['type'] == '携号转网开户':
            b_work_turn.append(i['work'])
for m in b_work:
    for n in points:
        if m == n['type']:
            b_point += n['cent']

for i in all_data:
    if i['name'] == '李娜':
        c_work.append(i['work'])
        if i['num'] == '1000':
            c_work_num.append(i['work'])
        if i['type'] == '携号转网开户':
            c_work_turn.append(i['work'])
for m in c_work:
    for n in points:
        if m == n['type']:
            c_point += n['cent']

for i in all_data:
    if i['name'] == '李荣':
        c_work.append(i['work'])
        if i['num'] == '1000':
            d_work_num.append(i['work'])
        if i['type'] == '携号转网开户':
            d_work_turn.append(i['work'])
for m in d_work:
    for n in points:
        if m == n['type']:
            d_point += n['cent']

tem_excel = xlrd.open_workbook('模板.xls', formatting_info=True)
tem_sheet = tem_excel.sheet_by_index(0)

new_excel = copy(tem_excel)
new_sheet = new_excel.get_sheet(0)

style = xlwt.XFStyle()

new_sheet.write(1, 1, a_work.count(active[2]))
new_sheet.write(1, 2, len(a_work_turn))
new_sheet.write(1, 3, len(a_work_num))
new_sheet.write(1, 4, a_point)
new_sheet.write(2, 1, b_work.count(s))
new_sheet.write(2, 2, len(b_work_turn))
new_sheet.write(2, 3, len(b_work_num))
new_sheet.write(2, 4, b_point)
new_sheet.write(3, 1, c_work.count(s))
new_sheet.write(3, 2, len(c_work_turn))
new_sheet.write(3, 3, len(c_work_num))
new_sheet.write(3, 4, c_point)
new_sheet.write(4, 1, d_work.count(s))
new_sheet.write(4, 2, len(d_work_turn))
new_sheet.write(4, 3, len(d_work_num))
new_sheet.write(4, 4, d_point)

new_excel.save('统计表.xls')
