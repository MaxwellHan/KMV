import xlrd
import xlwt
from scipy.stats import linregress
import scipy
import numpy as np
import statsmodels.api as sm
import warnings

warnings.filterwarnings("ignore") #屏蔽警告

#输入excel
inbook = xlrd.open_workbook('实证ZW\\beta7_qr.xlsx')
#输出excel
outbook = xlwt.Workbook()
sheet_out = outbook.add_sheet('sheet1')
out_book_path = 'result_beta_7_qr.xls'   #输出路径

#读取表格
sheet_names = inbook.sheet_names()
sheet_name = sheet_names[0]
sheet2 = inbook.sheet_by_name(sheet_name)

rowsnum = len(sheet2.col_values(0))#总的行数

#超参数
n_par = 4  #变量的个数
n_xpar = 0 #独立变量的个数，每个公司独有的一维变量
Reg_Method = 1 #回归类型 0是最小二乘线性回归， 1是分位数回归
y_nums = 111 #公司数

row = sheet2.row_values(1) #第二行
y_names = row[n_par+n_xpar*y_nums+1:n_par+n_xpar*y_nums+1+y_nums]  #变量名

#写入第一行标题
if Reg_Method == 0:
    for i in range(n_par+n_xpar+1):
        sheet_out.write(0, i + 1, '参数_' + str(i+1))
        sheet_out.write(0, i + 1 + (n_par+n_xpar+1), 't值_' + str(i + 1))
        sheet_out.write(0, i + 1 + 2*(n_par+n_xpar+1), 'p值_' + str(i + 1))
elif Reg_Method == 1:
    for i in range(n_par+1):
        sheet_out.write(0, i + 1, 'q.25_参数_' + str(i+1))
        sheet_out.write(0, i + 1 + (n_par+n_xpar+1), 'q.25_t值_' + str(i + 1))
        sheet_out.write(0, i + 1 + 2*(n_par+n_xpar+1), 'q.25_p值_' + str(i + 1))
        sheet_out.write(0, i + 1 + 3*(n_par+n_xpar+1), 'q.5_参数_' + str(i + 1))
        sheet_out.write(0, i + 1 + 4*(n_par+n_xpar + 1), 'q.5_t值_' + str(i + 1))
        sheet_out.write(0, i + 1 + 5 * (n_par+n_xpar + 1), 'q.5_p值_' + str(i + 1))
        sheet_out.write(0, i + 1 + 6 * (n_par+n_xpar + 1), 'q.75_参数_' + str(i + 1))
        sheet_out.write(0, i + 1 + 7*(n_par+n_xpar + 1), 'q.75_t值_' + str(i + 1))
        sheet_out.write(0, i + 1 + 8 * (n_par+n_xpar + 1), 'q.75_p值_' + str(i + 1))

#把名字写入输出表格
for i,name in enumerate(y_names):
    sheet_out.write(i+1,0,name)

#解析自变量x，放入x_mat中，x_mat是一个矩阵，大小为 样本数*(参数个数+1) 最后一维为常数 恒为1
#解析共有的自变量
x_mat = np.ones((rowsnum-2,n_par+n_xpar+1))
for i in range(2,rowsnum):
    row = sheet2.row_values(i)
    x_mat[i-2,:n_par] = np.float64(np.array(row[1:1+n_par]))

y = np.zeros(rowsnum-2)  #因变量

for i in range(len(y_names)):
# for i in range(1):
    loc = 1+n_par+i
    #解析y值
    for j in range(2,rowsnum):
       row = sheet2.row_values(j)
       y[j-2] = float(row[loc + n_xpar*y_nums])
    #解析每个公司独有的变量
    for j in range(2, rowsnum):
        row = sheet2.row_values(j)
        x_mat[j-2,n_par:n_par+n_xpar] = np.float64(np.array(row[loc+i*n_xpar:loc+(i+1)*n_xpar]))

    if Reg_Method == 0:
        est = sm.OLS(y, x_mat)
        res = est.fit()
        for k in range(n_par+n_xpar+1):
            sheet_out.write(i + 1, k + 1, res.params[k])
            sheet_out.write(i + 1, k + 1 + (n_par+n_xpar+1), res.tvalues[k])
            sheet_out.write(i + 1, k + 1 + 2*(n_par+n_xpar+1), res.pvalues[k])
    elif Reg_Method == 1:
        est_q = sm.QuantReg(y,x_mat)
        res_q_25 = est_q.fit(q=0.25)
        res_q_50 = est_q.fit(q=0.5)
        res_q_75 = est_q.fit(q=0.75)
        for k in range(n_par+1):
            sheet_out.write(i + 1, k + 1, res_q_25.params[k])
            sheet_out.write(i + 1, k + 1 + (n_par+n_xpar+1), res_q_25.tvalues[k])
            sheet_out.write(i + 1, k + 1 + 2*(n_par+n_xpar+1), res_q_25.pvalues[k])
            sheet_out.write(i + 1, k + 1 + 3*(n_par+n_xpar+1), res_q_50.params[k])
            sheet_out.write(i + 1, k + 1 + 4*(n_par+n_xpar+1), res_q_50.tvalues[k])
            sheet_out.write(i + 1, k + 1 + 5*(n_par+n_xpar+1), res_q_50.pvalues[k])
            sheet_out.write(i + 1, k + 1 + 6*(n_par+n_xpar+1), res_q_75.params[k])
            sheet_out.write(i + 1, k + 1 + 7*(n_par+n_xpar+1), res_q_75.tvalues[k])
            sheet_out.write(i + 1, k + 1 + 8*(n_par+n_xpar+1), res_q_75.pvalues[k])
    print(i+1,"/",len(y_names))

#保存
outbook.save(out_book_path)
