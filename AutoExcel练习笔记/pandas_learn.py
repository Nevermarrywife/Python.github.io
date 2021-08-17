import pandas   as pd
import numpy    as np

# s = pd.Series([1,3,4,np.nan,6,8])
#
# print(s)

data = pd.read_excel("3.xlsx")
data.head()

result1 = pd.pivot_table(data,index='营业厅名称',columns='操作备注',values='操作对象',aggfunc=np.count_nonzero)
result1.head()
print(result1)