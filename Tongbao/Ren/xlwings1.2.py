import xlwings as xw

"读取数据"
app = xw.App(visible=True, add_book=False)
app.display_alerts = False
app.screen_updating = False

data_2886 = []
data_kd = []
data_v = []
data_hcy = []
data_2816 = []

# wb_2886 = app.books.open('1.xls')  # 打开Excel
# sht_2886 = wb_2886.sheets['Sheet1']  # 指定工作表
# # print(sht_2886.range(5,2).value)  #读取相应单元格内容
# info_2886 = sht_2886.used_range  # 有效行
# # print(info.last_cell.row)   #总行数
# for n in range(4,40000): # 遍历2886，提取数据
#     name = sht_2886.range(n, 5).value
#     county = sht_2886.range(n, 2).value
#     channelname = sht_2886.range(n, 3).value
#     work = sht_2886.range(n, 12).value
#     num = sht_2886.range(n, 6).value
#     type = sht_2886.range(n, 7).value
#     data = {'name': name, 'work': work, 'num': num, 'type': type, 'channelname': channelname,'county':county}
#     data_2886.append(data)
# print("2886提取成功。。。")
# wb_2886.close()  # 关闭报表
# # app.quit()  # 关闭应用

# wb_kd = app.books.open('装机工单.xlsx')
# sht_kd = wb_kd.sheets['装机工单']
# info_kd = sht_kd.used_range
# for n in range(2, info_kd.last_cell.row):  #遍历宽带
#     type = sht_kd.range(n, 3).value
#     work = sht_kd.range(n, 5).value
#     name = sht_kd.range(n, 17).value
#     channelname = sht_kd.range(n, 18).value
#     data = {'name': name, 'work': work, 'type': type, 'channelname': channelname}
#     data_kd.append(data)
# print('宽带提取成功。。。')
# wb_kd.close()  # 关闭报表
# # app.quit()  # 关闭应用

periods = [{'period':'jlakEk','name':'李荣'},
           {'period':'jlakNO','name':'闫晓晶'},
           {'period':'jlak47','name':'石慧慧'},
           {'period':'jlaksi','name':'谢爱梅'},
           {'period':'jlakGK','name':'刘伟妮'},
           {'period':'jlakY6','name':'张文娜'},
           {'period':'jlakct','name':'张敏'},
           {'period':'jlakqd','name':'梁瑞丽'},
           {'period':'jlakxo','name':'赵晓丽'},
           {'period':'jlakA4','name':'郝晓婷'},
           {'period':'jlla7X','name':'李艳苗'},
           {'period':'jllaAJ','name':'柳步楠'},
           {'period':'jllaB3','name':'王莎莎'},
           {'period':'jlla9p','name':'方冰'},
           {'period':'jllaer','name':'张慧芳'},
           {'period':'jllape','name':'李娜'},
           {'period':'jllaCW','name':'姬慧婷'},
           {'period':'jlgk8T','name':'贾洋洋'},
           {'period':'jlgkZi','name':'郝晓霞'},
           {'period':'jlgkiv','name':'解晓娜'},
           {'period':'jlgkfu','name':'屈丽琴'},
           {'period':'jlgkbF','name':'郭娜'}]
countys = ['盐湖','芮城','平陆','临猗','万荣','河津','稷山','垣曲','绛县','闻喜','新绛','永济','夏县']
countys_yc = ['盐湖营业部','芮城分公司','平陆分公司','临猗分公司','万荣分公司','河津分公司','稷山分公司','垣曲分公司','绛县分公司','闻喜分公司','新绛分公司','永济分公司','夏县分公司']
# wb_yj = app.books.open('2816.xls')
# sht_yj = wb_yj.sheets['Sheet1']
# info_yj = sht_yj.used_range
# for n in range(4, info_yj.last_cell.row):  #遍历2816
#         work = sht_yj.range(n, 4).value
#         period = sht_yj.range(n, 2).value
#         channelname = sht_yj.range(n, 1).value
#         for m in countys:
#             if channelname.__contains__(m):
#                 county = m
#         data = {'period': period,'work': work,'channelname': channelname,'county':county}
#         data_2816.append(data)
# print('2816提取成功。。。')
# wb_yj.close()  # 关闭报表
# # app.quit()  # 关闭应用

wb_v = app.books.open('V.xls')
sht_v = wb_v.sheets['Sheet1']
info_v = sht_v.used_range
for n in range(4, info_v.last_cell.row):  #遍历V网
    type = sht_v.range(n, 10).value
    name = sht_v.range(n, 4).value
    channelname = sht_v.range(n, 2).value
    data = {'name': name,'type': type,'channelname': channelname}
    data_v.append(data)
print('V网提取成功。。。')
wb_v.close()  # 关闭报表
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
        fanghao.append(i['county'])
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

for i in data_2816:
    if i['work'] != None and i['work'] != '2020年“和卫士”（电子学生证）全省统一营销活动_学生参加(GK 309)':
        yingjian.append(i['period'])
        yingjian.append(i['county'])
        yingjian.append(i['channelname'])
for n in yingjian:
    for m in periods:
        if n == m['period']:
            print(n)
            print(m)
            yingjian.append(m['name'])
print('硬件筛选成功。。。。。')

for i in data_v:
    if i['type'] == '增加':
        netv.append(i['name'])
        netv.append(i['channelname'])
print('V网筛选成功。。。。。')

"写入数据"
app = xw.App(visible=True, add_book=False)
app.display_alerts = False
app.screen_updating = False

wb_tb = app.books.open('6月通报专用.xlsx')
sht_tb_yj = wb_tb.sheets['永济通报']
sht_tb_qx = wb_tb.sheets['区县通报']

for i in range(2,len(persons)+1):
    sht_tb_yj.range((i,8)).value = fanghao.count(persons[i-2])
    sht_tb_yj.range((i,4)).value = kuandai.count(persons[i-2])
    sht_tb_yj.range((i,6)).value = iptv.count(persons[i-2])
    sht_tb_yj.range((i,10)).value = xiezhuan.count(persons[i-2])
    sht_tb_yj.range((i,14)).value = netv.count(persons[i-2])
    sht_tb_yj.range((i, 16)).value = yingjian.count(persons[i-2])
for i in range(2,len(countys)+2):
    sht_tb_qx.range((i,10)).value = yingjian.count(countys[i-2])
    sht_tb_qx.range((i,2)).value = fanghao.count(countys_yc[i - 2])
print("数据写入成功!")

wb_tb.save()  # 保存报表
wb_tb.close()  # 关闭报表
app.quit()  # 关闭应用
