import xlrd

data = xlrd.open_workbook('test1.xlsx')
sheet = data.sheet_by_index(0)
questlist = []

class Question:
    pass

for i in range(sheet.nrows):
    if i>2:
        obj = Question()
        obj.cllss = sheet.cell(i,0).value
        obj.names = sheet.cell(i,1).value
        obj.mals = sheet.cell(i,2).value
        obj.ages = sheet.cell(i,3).value
        obj.scores = sheet.cell(i,4).value
        questlist.append(obj)
print(questlist[2])







