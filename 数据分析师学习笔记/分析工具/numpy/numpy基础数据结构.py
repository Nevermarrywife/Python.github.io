import numpy as np

"numpy数组的基本属性"
ar = np.array([[[1,2,3,4,5,6],[2,6,4,5,9],[2,6,4,5,9]],[1,2,3]],dtype=object)

print(ar)  #输出数组，数据中间没有逗号
print(ar.ndim)  #输出数组维度S的个数（维度）
print(ar.shape) #数组的行列
print(ar.size)  #数组的行列乘积·

"数组的十种生成模式"
ar1 = np.array(range(10))
ar2 = np.arange(10)
ar3 = np.array([[1,2,3],[4,5,6],list(range(3))])
ar4 = np.random.rand(10).reshape(2,5)   #随机生成十个0-1的数字，并按照2行5列排序
ar5 = np.arange(3.0,10,2)  #返回3.0-10.0，步长为2
ar6 = np.linspace(10,20,num=21,endpoint=True) #返回10-20之间，个数为21的均匀间隔样本,endpoint为False的话，返回就不包含20了
ar7 = np.linspace(10,20,num=21,retstep=True)   #retstep为True时，返回10-20的元组，并带步长
ar8 = np.zeros((3,4),np.int_)   #返回0填充的固定格式数组
ar9 = np.zeros_like(ar3)    #返回格式和ar3一样的被0填充的数组
ar10 = np.eye(5)    #创建5*5，中间都为1的数组
print(ar1)
print(ar2)
print(ar3)
print(ar4)
print(ar5)
print(ar6)
print(ar7)
print(ar8)
print(ar9)
print(ar10)