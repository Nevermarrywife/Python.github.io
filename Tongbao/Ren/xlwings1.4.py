import xlwings as xw
import os

"读取数据"
app = xw.App(visible=False, add_book=False)
app.display_alerts = False
app.screen_updating = False

data_2886 = []
data_kd = []
data_v = []
data_hcy = []
data_2816 = []

if os.path.exists('2886.xls'):
    wb_2886 = app.books.open('2886.xls')  # 打开Excel
    sht_2886 = wb_2886.sheets['Sheet1']  # 指定工作表
    # print(sht_2886.range(5,2).value)  #读取相应单元格内容
    info_2886 = sht_2886.used_range  # 有效行
    # print(info.last_cell.row)   #总行数
    for n in range(4,info_2886.last_cell.row): # 遍历2886，提取数据
        data_2886.append(sht_2886.range("B" + str(n) + ":" + "N" + str(n)).value)
    print("2886提取成功。。。")
    wb_2886.close()  # 关闭报表
# app.quit()  # 关闭应用

if os.path.exists('装机工单.xlsx'):
    wb_kd = app.books.open('装机工单.xlsx')
    sht_kd = wb_kd.sheets['装机工单']
    info_kd = sht_kd.used_range
    for n in range(2, info_kd.last_cell.row+1):  #遍历宽带
        data_kd.append(sht_kd.range("A" + str(n) + ":" + "Y" + str(n)).value)
    print('宽带提取成功。。。')
    wb_kd.close()  # 关闭报表
    # app.quit()  # 关闭应用

periods = [{'period': 'jlakEk', 'name': '李荣', 'hcy': 0},
           {'period': 'jlakNO', 'name': '闫晓晶', 'hcy': 0},
           {'period': 'jlak47', 'name': '石慧慧', 'hcy': 0},
           {'period': 'jlaksi', 'name': '谢爱梅', 'hcy': 0},
           {'period': 'jlakGK', 'name': '刘伟妮', 'hcy': 0},
           {'period': 'jlakY6', 'name': '张文娜', 'hcy': 0},
           {'period': 'jlakct', 'name': '张敏', 'hcy': 0},
           {'period': 'jlakqd', 'name': '梁瑞丽', 'hcy': 0},
           {'period': 'jlakxo', 'name': '赵晓丽', 'hcy': 0},
           {'period': 'jlakA4', 'name': '郝晓婷', 'hcy': 0},
           {'period': 'jlla7X', 'name': '李艳苗', 'hcy': 0},
           {'period': 'jllaAJ', 'name': '柳步楠', 'hcy': 0},
           {'period': 'jllaB3', 'name': '王莎莎', 'hcy': 0},
           {'period': 'jlla9p', 'name': '方冰', 'hcy': 0},
           {'period': 'jllaer', 'name': '张慧芳', 'hcy': 0},
           {'period': 'jllape', 'name': '李娜', 'hcy': 0},
           {'period': 'jllaCW', 'name': '姬慧婷', 'hcy': 0},
           {'period': 'jlgk8T', 'name': '贾洋洋', 'hcy': 0},
           {'period': 'jlgkZi', 'name': '郝晓霞', 'hcy': 0},
           {'period': 'jlgkiv', 'name': '解晓娜', 'hcy': 0},
           {'period': 'jlgkfu', 'name': '屈丽琴', 'hcy': 0},
           {'period': 'jlgkbF', 'name': '郭娜', 'hcy': 0},
           {'period': 'jlpv3C', 'name': '谷晶', 'hcy': 0},
           {'period': 'jlpvHn', 'name': '谭莹', 'hcy': 0},
           {'period': 'jlpvsZ', 'name': '樊海婷', 'hcy': 0},
           {'period': 'jlpv9F', 'name': '吴茜茜', 'hcy': 0},
           {'period': 'jlpwnN', 'name': '邓志强', 'hcy': 0},
           {'period': 's', 'name': '程小文', 'hcy': 0},
           {'period': 'jlpwLw', 'name': '王洁煊', 'hcy': 0},
           {'period': 'jlpwFB', 'name': '寻晓慧', 'hcy': 0},
           {'period': 's', 'name': '张冰', 'hcy': 0},
           {'period': 'jlpxJN', 'name': '武婷', 'hcy': 0},
           {'period': 's', 'name': '杨森', 'hcy': 0},
           {'period': 'jlpxoy', 'name': '赵娜', 'hcy': 0},
           {'period': 'jlpx52', 'name': '王园祺', 'hcy': 0}]
countys = ['盐湖', '芮城', '平陆', '临猗', '万荣', '河津', '稷山', '垣曲', '绛县', '闻喜', '新绛', '永济', '夏县']
countys_yc = ['盐湖营业部', '芮城分公司', '平陆分公司', '临猗分公司', '万荣分公司', '河津分公司', '稷山分公司', '垣曲分公司', '绛县分公司', '闻喜分公司', '新绛分公司',
              '永济分公司', '夏县分公司']

if os.path.exists('2816.xls'):
    wb_yj = app.books.open('2816.xls')
    sht_yj = wb_yj.sheets['Sheet1']
    info_yj = sht_yj.used_range
    for n in range(4, info_yj.last_cell.row+1):  #遍历2816
            work = sht_yj.range(n, 4).value
            period = sht_yj.range(n, 2).value
            channelname = sht_yj.range(n, 1).value
            for m in countys:
                if channelname.__contains__(m):
                    county = m
            data = {'period': period,'work': work,'channelname': channelname,'county':county}
            data_2816.append(data)
    print('2816提取成功。。。')
    wb_yj.close()  # 关闭报表
    # app.quit()  # 关闭应用

if os.path.exists('hcy.xlsx'):
    wb_hcy = app.books.open('hcy.xlsx')
    sht_hcy = wb_hcy.sheets['Sheet1']
    info_hcy = sht_hcy.used_range
    for n in range(3, info_hcy.last_cell.row + 1):  # 遍历宽带
        data_hcy.append(sht_hcy.range("C" + str(n) + ":" + "X" + str(n)).value)
    print('和彩云提取成功。。。')
    wb_hcy.close()  # 关闭报表
    # app.quit()  # 关闭应用

if os.path.exists('280n.xls'):
    wb_v = app.books.open('280n.xls')
    sht_v = wb_v.sheets['Sheet1']
    info_v = sht_v.used_range
    for n in range(4, info_v.last_cell.row+1):  #遍历V网
        channelname = sht_v.range(n, 1).value
        person = sht_v.range(n, 3).value
        work = sht_v.range(n, 10).value
        for m in countys:
            if channelname.__contains__(m):
                county = m
        data = {'person': person,'work': work,'channelname': channelname,'county':county}
        data_v.append(data)
    print('V网提取成功。。。')
    wb_v.close()  # 关闭报表
    app.quit()  # 关闭应用


"筛选数据"
fanghao = []
fanghao_yj = []
ganen = []
tehuibao = []
jifen = []
kuandai = []
iptv = []
xiezhuan = []
xiezhuan_yj = []
netv = []
netv_yj = []
yingjian = []
hcy = []

persons = ('李荣', '闫晓晶', '石慧慧', '谢爱梅', '刘伟妮', '张文娜', '张敏', '梁瑞丽', '赵晓丽', '郝晓婷', '李艳苗', '柳步楠',
           '王莎莎', '方冰', '张慧芳', '李娜', '姬慧婷', '贾洋洋', '郝晓霞', '解晓娜', '屈丽琴', '郭娜', '运城永济城东关铝手机专卖店',
           '运城永济城区郭李手机专卖店', '运城永济于乡镇清华手机专卖店', '运城永济城区迎新手机专卖店', '运城永济于乡镇于乡二部手机专卖店',
           '运城永济城东侯孟手机专卖店', '运城永济城区银杏手机专卖店', '运城永济卿头镇董村手机专卖店', '运城永济卿头镇许家营手机专卖店',
           '运城永济栲栳镇韩村手机专卖店', '运城永济栲栳镇栲栳手机专卖店', '运城永济栲栳镇缄庄手机专卖店', '运城永济张营镇张营手机专卖店',
           '运城永济开张镇黄营手机专卖店', '运城永济城区北郊手机专卖店', '运城永济开张镇开张手机专卖店', '运城永济城区电机手机专卖店',
           '运城永济城区樱花手机专卖店', '运城永济韩阳镇韩阳手机专卖店', '运城永济城西七社手机专卖店', '运城永济城区晋通手机专卖店',
           '运城永济蒲州镇文学手机专卖店', '运城永济城区赵柏手机专卖店', '运城永济城区永纺手机专卖店', '运城永济城区四冯手机专卖店', '运城永济蒲州镇西厢手机专卖店',
           '谷晶', '谭莹', '樊海婷', '吴茜茜', '邓志强', '程小文', '王洁煊', '寻晓慧', '张冰', '武婷', '杨森', '赵娜', '王园祺')

active_v = ('JT2130-运城二季度和飞信新增活动（3个月）', 'JT2131-运城二季度和飞信升档活动（3个月）',
            'JT2123-运城二季度和飞信升档活动（6个月）', 'JT2122-运城二季度和飞信升档活动（4个月）',
            'JT2120-运城二季度和飞信新增活动（4个月）', 'JT2121-运城二季度和飞信新增活动（6个月）')

for i in data_2886:
    if i[4] == '1000' or i[4] == '1379' or i[4] == '4696':
        fanghao.append(i[0])
        if i[0] == '永济分公司':
            fanghao_yj.append(i[3])
            fanghao_yj.append(i[1])
    if i[0] == '永济分公司' and i[4] == '4696':
        xiezhuan_yj.append(i[3])
        xiezhuan_yj.append(i[1])
    # for m in range(0, len(active_v)):  # 筛选1147和飞信
    #     if i[10] == active_v[m]:
    #         netv.append(i[0])
    #         if i[0] == '永济分公司':
    #             netv_yj.append(i[1])
    #             netv_yj.append(i[3])


for i in data_kd:
    if i[4] == '家庭宽带' or i[4] == '家庭宽带+IPTV':
        kuandai.append(i[16])
        kuandai.append(i[17])
    if i[4] == '基于宽带的IPTV加装' or i[4] == '家庭宽带+IPTV':
        iptv.append(i[16])
        iptv.append(i[17])

for i in data_2816:
    if i['work'] != None and i['work'] != '2020年“和卫士”（电子学生证）全省统一营销活动_学生参加(GK 309)':
        yingjian.append(i['period'])
        yingjian.append(i['county'])
        yingjian.append(i['channelname'])
for n in yingjian:
    for m in periods:
        if n == m['period']:
            yingjian.append(m['name'])

for n in data_hcy:
    for m in periods:
        if n[0] == m['name']:
            ln = list(filter(lambda n: isinstance(n, str), n))  # 剔除掉None，避免计数时算上
            m['hcy'] += len(ln) - 2

for i in data_v:
    if i['work'] == '增加':

print('V网筛选成功。。。。。')

"写入数据"
app = xw.App(visible=False, add_book=False)
app.display_alerts = False
app.screen_updating = False

wb_tb = app.books.open('6月通报专用.xlsx')
sht_tb_yj = wb_tb.sheets['永济通报']
sht_tb_qx = wb_tb.sheets['区县通报']

for i in range(2, len(persons) + 2):
    sht_tb_yj.range((i, 8)).value = fanghao_yj.count(persons[i - 2])
    sht_tb_yj.range((i, 4)).value = kuandai.count(persons[i - 2])
    sht_tb_yj.range((i, 6)).value = iptv.count(persons[i - 2])
    sht_tb_yj.range((i, 10)).value = xiezhuan_yj.count(persons[i - 2])
    sht_tb_yj.range((i, 14)).value = netv_yj.count(persons[i - 2])
    sht_tb_yj.range((i, 16)).value = yingjian.count(persons[i - 2])
for i in range(2, len(periods) + 2):  # 写入和彩云
    sht_tb_yj.range((i, 2)).value = periods[i - 2]['hcy']
for i in range(2, len(countys) + 2):
    sht_tb_qx.range((i, 10)).value = yingjian.count(countys[i - 2])
    sht_tb_qx.range((i, 2)).value = fanghao.count(countys_yc[i - 2])
    sht_tb_qx.range((i, 12)).value = netv.count(countys_yc[i - 2])
print("数据写入成功!")

wb_tb.save()  # 保存报表
wb_tb.close()  # 关闭报表
app.quit()  # 关闭应用
