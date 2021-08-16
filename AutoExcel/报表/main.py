import xlrd
import xlwt
from xlutils.copy import copy
import xlwings as xw

xls = xlrd.open_workbook('1.xls')
table = xls.sheet_by_index(0)

all_data = []
s = 0

for n in range(3, table.nrows):
    name = table.cell(n, 4).value
    channelname = table.cell(n,2).value
    work = table.cell(n, 11).value
    num = table.cell(n, 5).value
    type = table.cell(n, 6).value
    obj = table.cell(n,7).value
    data = {'name': name,'work': work, 'num': num, 'type': type,'obj':obj,'channelname':channelname}
    all_data.append(data)

fanghao = []
ganen = []
tehuibao = []
jifen = []
kuandai = []
xiezhuan = []
C = 0

k = 'SC2171-二季度家庭融合享优惠（副卡）'
g = 'SC2172-二季度家庭融合享优惠（家庭网）'
s = 'SC2187-运城移动客户感恩有礼回馈活动'

active = ['SC2171-二季度家庭融合享优惠（副卡）', 'SC2172-二季度家庭融合享优惠（家庭网）', 'SC2187-运城移动客户感恩有礼回馈活动']

points = [{'type': '2019积分兑换活动(3G全国流量)', 'cent': 1000},
          {'type': '2019积分兑换活动(9G全国流量)', 'cent': 3000},
          {'type': '2019年积分兑换直充包（流量特惠包1GB(24小时)）', 'cent': 5000}]

for i in all_data:
    if i['num'] == '1000' or i['num'] == '1379' or i['num'] == '4696':
        fanghao.append(i['name'])
        fanghao.append(i['channelname'])
    if i['num'] == '4696':
        xiezhuan.append(i['name'])
        xiezhuan.append(i['channelname'])

tem_excel = xlrd.open_workbook('6月通报专用.xls', formatting_info=True)
tem_sheet = tem_excel.sheet_by_index(7)

new_excel = copy(tem_excel)
new_sheet = new_excel.get_sheet(7)

new_sheet.write(1, 6, '张文娜')
new_sheet.write(1, 7, fanghao.count('张文娜'))
new_sheet.write(2, 6, '李娜')
new_sheet.write(2, 7, fanghao.count('李娜'))
new_sheet.write(3, 6, '运城永济韩阳镇韩阳手机专卖店')
new_sheet.write(3, 7, fanghao.count('运城永济韩阳镇韩阳手机专卖店'))
new_sheet.write(4, 6, '运城永济卿头镇许家营手机专卖店')
new_sheet.write(4, 7, fanghao.count('运城永济卿头镇许家营手机专卖店'))

new_excel.save('统计表.xls')