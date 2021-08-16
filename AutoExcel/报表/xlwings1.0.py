import xlwings as xw

"读取数据"
app = xw.App(visible=True, add_book=False)
app.display_alerts = True
app.screen_updating = True

wb_2886 = app.books.open('D:/ai/1.xls')  # 打开Excel
sht_2886 = wb_2886.sheets['Sheet1']  # 指定工作表
# print(sht_2886.range(5,2).value)  #读取相应单元格内容
info_2886 = sht_2886.used_range  # 有效行
# print(info.last_cell.row)   #总行数

wb_kd = app.books.open('D:/ai/装机工单.xlsx')
sht_kd = wb_kd.sheets['装机工单']
info_kd = sht_kd.used_range

wb_v = app.books.open('D:/ai/V.xls')
sht_v = wb_v.sheets['Sheet1']
info_v = sht_v.used_range

# wb_yj = app.books.open('2816.xls')
# sht_yj = wb_yj.sheets['Sheet1']
# info_yj = sht_yj.used_range

data_2886 = []
data_kd = []
data_v = []
data_hcy = []
data_2816 = []

for n in range(4, info_2886.last_cell.row):  # 遍历2886，提取数据
    name = sht_2886.range(n, 5).value
    channelname = sht_2886.range(n, 3).value
    work = sht_2886.range(n, 12).value
    num = sht_2886.range(n, 6).value
    type = sht_2886.range(n, 7).value
    obj = sht_2886.range(n, 8).value
    period = sht_2886.range(n, 4).value
    data = {'name': name, 'work': work, 'num': num, 'type': type, 'obj': obj, 'channelname': channelname,
            'period': period}
    data_2886.append(data)
print("2886提取成功。。。")

for n in range(2, info_kd.last_cell.row):  #遍历宽带
    type = sht_kd.range(n, 3).value
    work = sht_kd.range(n, 5).value
    name = sht_kd.range(n, 17).value
    channelname = sht_kd.range(n, 18).value
    data = {'name': name, 'work': work, 'type': type, 'channelname': channelname}
    data_kd.append(data)
print('宽带提取成功。。。')

for n in range(4, info_v.last_cell.row):  #遍历V网
    type = sht_v.range(n, 10).value
    name = sht_v.range(n, 4).value
    channelname = sht_v.range(n, 2).value
    data = {'name': name,'type': type,'channelname': channelname}
    data_v.append(data)
print('V网提取成功。。。')

# for n in range(3, info_yj.last_cell.row):  # 遍历2816，提取数据
#     channelname = sht_yj.range(n, 1).value
#     work = sht_yj.range(n, 4).value
#     period = sht_yj.range(n, 2).value
#     data = {'work': work,  'channelname': channelname,'period': period}
#     data_2816.append(data)

wb_2886.close()  # 关闭报表
wb_kd.close()  # 关闭报表
wb_v.close()  # 关闭报表
# wb_yj.close()  # 关闭报表
app.quit()  # 关闭应用

"筛选数据"
fanghao = []
ganen = []
tehuibao = []
jifen = []
kuandai = []
iptv = []
xiezhuan = []
netv = []
yingjian = []

k = 'SC2171-二季度家庭融合享优惠（副卡）'
g = 'SC2172-二季度家庭融合享优惠（家庭网）'
s = 'SC2187-运城移动客户感恩有礼回馈活动'

active = ['SC2171-二季度家庭融合享优惠（副卡）', 'SC2172-二季度家庭融合享优惠（家庭网）', 'SC2187-运城移动客户感恩有礼回馈活动']

points = [{'type': '2019积分兑换活动(3G全国流量)', 'cent': 1000},
          {'type': '2019积分兑换活动(9G全国流量)', 'cent': 3000},
          {'type': '2019年积分兑换直充包（流量特惠包1GB(24小时)）', 'cent': 5000}]

periods = [{'period':'jlakqd','name':'梁瑞丽'},{'period':'jlaksi','name':'谢爱梅'}]

persons = ('李荣','闫晓晶','石慧慧','谢爱梅','刘伟妮','张文娜','张敏','梁瑞丽','赵晓丽','郝晓婷','李艳苗','柳步楠',
           '王莎莎','方冰','张慧芳','李娜','姬慧婷','贾洋洋','郝晓霞','解晓娜','屈丽琴','郭娜','运城永济城东关铝手机专卖店',
           '运城永济城区郭李手机专卖店','运城永济于乡镇清华手机专卖店','运城永济城区迎新手机专卖店','运城永济于乡镇于乡二部手机专卖店',
           '运城永济城东侯孟手机专卖店','运城永济城区银杏手机专卖店','运城永济卿头镇董村手机专卖店','运城永济卿头镇许家营手机专卖店',
           '运城永济栲栳镇韩村手机专卖店','运城永济栲栳镇栲栳手机专卖店','运城永济栲栳镇缄庄手机专卖店','运城永济张营镇张营手机专卖店',
           '运城永济开张镇黄营手机专卖店','运城永济城区北郊手机专卖店','运城永济开张镇开张手机专卖店','运城永济城区电机手机专卖店',
           '运城永济城区樱花手机专卖店','运城永济韩阳镇韩阳手机专卖店','运城永济城西七社手机专卖店','运城永济城区晋通手机专卖店',
           '运城永济蒲州镇文学手机专卖店','运城永济城区赵柏手机专卖店','运城永济城区永纺手机专卖店','运城永济城区四冯手机专卖店','运城永济蒲州镇西厢手机专卖店')


for i in data_2886:
    if i['num'] == '1000' or i['num'] == '1379' or i['num'] == '4696':
        fanghao.append(i['name'])
        fanghao.append(i['channelname'])
    if i['num'] == '4696':
        xiezhuan.append(i['name'])
        xiezhuan.append(i['channelname'])
print('2886筛选成功。。。。。')

for i in data_kd:
    if i['work'] == '家庭宽带' or i['work'] == '家庭宽带+IPTV':
        kuandai.append(i['name'])
        kuandai.append(i['channelname'])
    if i['work'] == '基于宽带的IPTV加装' or i['work'] == '家庭宽带+IPTV':
        iptv.append(i['name'])
        iptv.append(i['channelname'])
print('宽带筛选成功。。。。。')

for i in data_v:
    if i['type'] == '增加':
        netv.append(i['name'])
        netv.append(i['channelname'])
print('V网筛选成功。。。。。')

# for i in data_2816:
#     for n in range(0,len(periods)):
#         if i['period'] == periods[n]['period'] and i['work'] == '智家新人礼':
#             yingjian.append(periods[n]['name'])

"写入数据"
app = xw.App(visible=True, add_book=False)
app.display_alerts = True
app.screen_updating = True

wb_tb = app.books.open('6月通报专用.xlsx')
sht_tb_yj = wb_tb.sheets['永济通报']

for i in range(2,len(persons)+1):
    # sht_tb_yj.range((i,7)).value = persons[i-2]
    sht_tb_yj.range((i,8)).value = fanghao.count(persons[i-2])
    # sht_tb_yj.range((i,3)).value = persons[i-2]
    sht_tb_yj.range((i,4)).value = kuandai.count(persons[i-2])
    # sht_tb_yj.range((i,5)).value = persons[i-2]
    sht_tb_yj.range((i,6)).value = iptv.count(persons[i-2])
    # sht_tb_yj.range((i,9)).value = persons[i-2]
    sht_tb_yj.range((i,10)).value = xiezhuan.count(persons[i-2])
    # sht_tb_yj.range((i,13)).value = persons[i-2]
    sht_tb_yj.range((i,14)).value = netv.count(persons[i-2])
print("写入成功")
"""
sht_tb_yj.range('G2').value = '方冰'
sht_tb_yj.range('H2').value = fanghao.count('方冰')
sht_tb_yj.range('C2').value = '方冰'
sht_tb_yj.range('D2').value = kuandai.count('方冰')
sht_tb_yj.range('E2').value = '方冰'
sht_tb_yj.range('F2').value = iptv.count('方冰')
sht_tb_yj.range('I2').value = '方冰'
sht_tb_yj.range('J2').value = xiezhuan.count('方冰')
sht_tb_yj.range('M2').value = '方冰'
sht_tb_yj.range('N2').value = netv.count('方冰')

sht_tb_yj.range('G3').value = '李荣'
sht_tb_yj.range('H3').value = fanghao.count('李荣')
sht_tb_yj.range('C3').value = '李荣'
sht_tb_yj.range('D3').value = kuandai.count('李荣')
sht_tb_yj.range('E3').value = '李荣'
sht_tb_yj.range('F3').value = iptv.count('李荣')
sht_tb_yj.range('I3').value = '李荣'
sht_tb_yj.range('J3').value = xiezhuan.count('李荣')
sht_tb_yj.range('M3').value = '李荣'
sht_tb_yj.range('N3').value = netv.count('李荣')

sht_tb_yj.range('G4').value = '闫晓晶'
sht_tb_yj.range('H4').value = fanghao.count('闫晓晶')
sht_tb_yj.range('C4').value = '闫晓晶'
sht_tb_yj.range('D4').value = kuandai.count('闫晓晶')
sht_tb_yj.range('E4').value = '闫晓晶'
sht_tb_yj.range('F4').value = iptv.count('闫晓晶')
sht_tb_yj.range('I4').value = '闫晓晶'
sht_tb_yj.range('J4').value = xiezhuan.count('闫晓晶')
sht_tb_yj.range('M4').value = '闫晓晶'
sht_tb_yj.range('N4').value = netv.count('闫晓晶')

sht_tb_yj.range('G5').value = '石慧慧'
sht_tb_yj.range('H5').value = fanghao.count('石慧慧')
sht_tb_yj.range('C5').value = '石慧慧'
sht_tb_yj.range('D5').value = kuandai.count('石慧慧')
sht_tb_yj.range('E5').value = '石慧慧'
sht_tb_yj.range('F5').value = iptv.count('石慧慧')
sht_tb_yj.range('I5').value = '石慧慧'
sht_tb_yj.range('J5').value = xiezhuan.count('石慧慧')
sht_tb_yj.range('M5').value = '石慧慧'
sht_tb_yj.range('N5').value = netv.count('石慧慧')

sht_tb_yj.range('G6').value = '谢爱梅'
sht_tb_yj.range('H6').value = fanghao.count('谢爱梅')
sht_tb_yj.range('C6').value = '谢爱梅'
sht_tb_yj.range('D6').value = kuandai.count('谢爱梅')
sht_tb_yj.range('E6').value = '谢爱梅'
sht_tb_yj.range('F6').value = iptv.count('谢爱梅')
sht_tb_yj.range('I6').value = '谢爱梅'
sht_tb_yj.range('J6').value = xiezhuan.count('谢爱梅')
sht_tb_yj.range('M6').value = '谢爱梅'
sht_tb_yj.range('N6').value = netv.count('谢爱梅')

sht_tb_yj.range('G7').value = '张文娜'
sht_tb_yj.range('H7').value = fanghao.count('张文娜')
sht_tb_yj.range('C7').value = '张文娜'
sht_tb_yj.range('D7').value = kuandai.count('张文娜')
sht_tb_yj.range('E7').value = '张文娜'
sht_tb_yj.range('F7').value = iptv.count('张文娜')
sht_tb_yj.range('I7').value = '张文娜'
sht_tb_yj.range('J7').value = xiezhuan.count('张文娜')
sht_tb_yj.range('M7').value = '张文娜'
sht_tb_yj.range('N7').value = netv.count('张文娜')

sht_tb_yj.range('G8').value = '张敏'
sht_tb_yj.range('H8').value = fanghao.count('张敏')
sht_tb_yj.range('C8').value = '张敏'
sht_tb_yj.range('D8').value = kuandai.count('张敏')
sht_tb_yj.range('E8').value = '张敏'
sht_tb_yj.range('F8').value = iptv.count('张敏')
sht_tb_yj.range('I8').value = '张敏'
sht_tb_yj.range('J8').value = xiezhuan.count('张敏')
sht_tb_yj.range('M8').value = '张敏'
sht_tb_yj.range('N8').value = netv.count('张敏')

sht_tb_yj.range('G9').value = '梁瑞丽'
sht_tb_yj.range('H9').value = fanghao.count('梁瑞丽')
sht_tb_yj.range('C9').value = '梁瑞丽'
sht_tb_yj.range('D9').value = kuandai.count('梁瑞丽')
sht_tb_yj.range('E9').value = '梁瑞丽'
sht_tb_yj.range('F9').value = iptv.count('梁瑞丽')
sht_tb_yj.range('I9').value = '梁瑞丽'
sht_tb_yj.range('J9').value = xiezhuan.count('梁瑞丽')
sht_tb_yj.range('M9').value = '梁瑞丽'
sht_tb_yj.range('N9').value = netv.count('梁瑞丽')

sht_tb_yj.range('G10').value = '赵晓丽'
sht_tb_yj.range('H10').value = fanghao.count('赵晓丽')
sht_tb_yj.range('C10').value = '赵晓丽'
sht_tb_yj.range('D10').value = kuandai.count('赵晓丽')
sht_tb_yj.range('E10').value = '赵晓丽'
sht_tb_yj.range('F10').value = iptv.count('赵晓丽')
sht_tb_yj.range('I10').value = '赵晓丽'
sht_tb_yj.range('J10').value = xiezhuan.count('赵晓丽')
sht_tb_yj.range('M10').value = '赵晓丽'
sht_tb_yj.range('N10').value = netv.count('赵晓丽')

sht_tb_yj.range('G11').value = '李艳苗'
sht_tb_yj.range('H11').value = fanghao.count('李艳苗')
sht_tb_yj.range('C11').value = '李艳苗'
sht_tb_yj.range('D11').value = kuandai.count('李艳苗')
sht_tb_yj.range('E11').value = '李艳苗'
sht_tb_yj.range('F11').value = iptv.count('李艳苗')
sht_tb_yj.range('I11').value = '李艳苗'
sht_tb_yj.range('J11').value = xiezhuan.count('李艳苗')
sht_tb_yj.range('M11').value = '李艳苗'
sht_tb_yj.range('N11').value = netv.count('李艳苗')

sht_tb_yj.range('G12').value = '柳步楠'
sht_tb_yj.range('H12').value = fanghao.count('柳步楠')
sht_tb_yj.range('C12').value = '柳步楠'
sht_tb_yj.range('D12').value = kuandai.count('柳步楠')
sht_tb_yj.range('E12').value = '柳步楠'
sht_tb_yj.range('F12').value = iptv.count('柳步楠')
sht_tb_yj.range('I12').value = '柳步楠'
sht_tb_yj.range('J12').value = xiezhuan.count('柳步楠')
sht_tb_yj.range('M12').value = '柳步楠'
sht_tb_yj.range('N12').value = netv.count('柳步楠')

sht_tb_yj.range('G13').value = '张慧芳'
sht_tb_yj.range('H13').value = fanghao.count('张慧芳')
sht_tb_yj.range('C13').value = '张慧芳'
sht_tb_yj.range('D13').value = kuandai.count('张慧芳')
sht_tb_yj.range('E13').value = '张慧芳'
sht_tb_yj.range('F13').value = iptv.count('张慧芳')
sht_tb_yj.range('I13').value = '张慧芳'
sht_tb_yj.range('J13').value = xiezhuan.count('张慧芳')
sht_tb_yj.range('M13').value = '张慧芳'
sht_tb_yj.range('N13').value = netv.count('张慧芳')

sht_tb_yj.range('G14').value = '李娜'
sht_tb_yj.range('H14').value = fanghao.count('李娜')
sht_tb_yj.range('C14').value = '李娜'
sht_tb_yj.range('D14').value = kuandai.count('李娜')
sht_tb_yj.range('E14').value = '李娜'
sht_tb_yj.range('F14').value = iptv.count('李娜')
sht_tb_yj.range('I14').value = '李娜'
sht_tb_yj.range('J14').value = xiezhuan.count('李娜')
sht_tb_yj.range('M14').value = '李娜'
sht_tb_yj.range('N14').value = netv.count('李娜')

sht_tb_yj.range('G15').value = '贾洋洋'
sht_tb_yj.range('H15').value = fanghao.count('贾洋洋')
sht_tb_yj.range('C15').value = '贾洋洋'
sht_tb_yj.range('D15').value = kuandai.count('贾洋洋')
sht_tb_yj.range('E15').value = '贾洋洋'
sht_tb_yj.range('F15').value = iptv.count('贾洋洋')
sht_tb_yj.range('I15').value = '贾洋洋'
sht_tb_yj.range('J15').value = xiezhuan.count('贾洋洋')
sht_tb_yj.range('M15').value = '贾洋洋'
sht_tb_yj.range('N15').value = netv.count('贾洋洋')

sht_tb_yj.range('G16').value = '郝晓霞'
sht_tb_yj.range('H16').value = fanghao.count('郝晓霞')
sht_tb_yj.range('C16').value = '郝晓霞'
sht_tb_yj.range('D16').value = kuandai.count('郝晓霞')
sht_tb_yj.range('E16').value = '郝晓霞'
sht_tb_yj.range('F16').value = iptv.count('郝晓霞')
sht_tb_yj.range('I16').value = '郝晓霞'
sht_tb_yj.range('J16').value = xiezhuan.count('郝晓霞')
sht_tb_yj.range('M16').value = '郝晓霞'
sht_tb_yj.range('N16').value = netv.count('郝晓霞')

sht_tb_yj.range('G17').value = '解晓娜'
sht_tb_yj.range('H17').value = fanghao.count('解晓娜')
sht_tb_yj.range('C17').value = '解晓娜'
sht_tb_yj.range('D17').value = kuandai.count('解晓娜')
sht_tb_yj.range('E17').value = '解晓娜'
sht_tb_yj.range('F17').value = iptv.count('解晓娜')
sht_tb_yj.range('I17').value = '解晓娜'
sht_tb_yj.range('J17').value = xiezhuan.count('解晓娜')
sht_tb_yj.range('M17').value = '解晓娜'
sht_tb_yj.range('N17').value = netv.count('解晓娜')

sht_tb_yj.range('G18').value = '屈丽琴'
sht_tb_yj.range('H18').value = fanghao.count('屈丽琴')
sht_tb_yj.range('C18').value = '屈丽琴'
sht_tb_yj.range('D18').value = kuandai.count('屈丽琴')
sht_tb_yj.range('E18').value = '屈丽琴'
sht_tb_yj.range('F18').value = iptv.count('屈丽琴')
sht_tb_yj.range('I18').value = '屈丽琴'
sht_tb_yj.range('J18').value = xiezhuan.count('屈丽琴')
sht_tb_yj.range('M18').value = '屈丽琴'
sht_tb_yj.range('N18').value = netv.count('屈丽琴')
print('写入成功。。。。。')
"""

wb_tb.save()  # 保存报表
wb_tb.close()  # 关闭报表
app.quit()  # 关闭应用
