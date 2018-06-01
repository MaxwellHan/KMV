import xlrd
import xlwt
from KMVfun import *
import numpy as np

#输入excel
inbook = xlrd.open_workbook('data1_y.xlsx')
#输出excel
outbook = xlwt.Workbook()
sheet_out = outbook.add_sheet('sheet1')

#超参数，设置债券要素的列数、股价的月数
n_Yaosu = 5  #债券要素的列数
n_Month = 3  #月份数
n_sigma = 3  #sigma的数量


sheet_names = inbook.sheet_names()
sheet_name = sheet_names[1]
sheet2 = inbook.sheet_by_name(sheet_name)
rowsnum = len(sheet2.col_values(0))
#前两行
rows = sheet2.row_values(0)  #读取第一行，标题行
[sheet_out.write(0,i,s) for i,s in enumerate(rows[0:n_Yaosu])]
[sheet_out.write(0,i,'公司资产价值') for i in range(n_Yaosu,n_Yaosu+n_Month*2,2)]
[sheet_out.write(0,i+1,'公司资产波动率') for i in range(n_Yaosu,n_Yaosu+n_Month*2,2)]
rows = sheet2.row_values(1) #第一行
[sheet_out.write(1,i,s) for i,s in enumerate(rows[0:n_Yaosu])]
[sheet_out.write(1,n_Yaosu+2*i,s) for i,s in enumerate(rows[n_Yaosu:n_Yaosu+n_Month])]
[sheet_out.write(1,n_Yaosu+2*i+1,s) for i,s in enumerate(rows[n_Yaosu:n_Yaosu+n_Month])]

for i in range(2,rowsnum):
    rows = sheet2.row_values(i)
    data_info = rows[0:n_Yaosu] #数据冗余信息
    print(data_info)
    [sheet_out.write(i, k, s) for k, s in enumerate(data_info)] #填入前n_Yaosu行
    #读取无风险利率
    if i==2:
        rate = rows[-n_Month:]
    bg = n_Yaosu
    ed = n_Yaosu+n_Month
    S_t = rows[bg:ed]    #读取股权价值 48列

    bg = ed
    ed += n_sigma
    if n_sigma == 1:  #一个波动率
        sigma = rows[bg]
    else:
        sigma = rows[bg:ed] #读取月化波动 48列
    bg = ed
    ed += n_Month
    Debt = rows[bg:ed] #

    bg = ed
    ed += n_Month
    tua = rows[bg]
    #开始逐月计算
    for j in range(n_Month):
        r = rate[j]*0.01  #无风险利率
        T = tua        #
        D = Debt[j]  #债务
        if n_sigma == 1:
            EquityTheta = sigma*0.01
        else:
            EquityTheta = sigma[j] * 0.01 #公司的股权价值波动率

        E = S_t[j]  #公司的股权价值
        try:
            res = KMVOptSearch(E, D, r, T, EquityTheta)
        except :
            print(E,D,r,T,EquityTheta)
        #写出到excel
        sheet_out.write(i,2*j+n_Yaosu,res[0]*E)
        sheet_out.write(i,2*j+n_Yaosu+1,res[1]*100) #乘以100
    print(i-1,"/",rowsnum-2)

outbook.save("result1_y.xls")