import scipy as sc
import scipy.stats as stats
from scipy.optimize import fsolve
import warnings
warnings.filterwarnings('ignore')
def KMVfun(EtoD,r,T,EquityTheta,x):
    d1 = (sc.log(x[0] * EtoD) + (r + 0.5 * x[1]**2) * T) / (x[1] * sc.sqrt(T))
    d2 = d1 - x[1] * sc.sqrt(T)
    return [x[0] * stats.norm.cdf(d1) - sc.exp(-r*T)*stats.norm.cdf(d2)/EtoD -1,
            stats.norm.cdf(d1)*x[0]*x[1]-EquityTheta]

def KMVOptSearch(E,D,r,T,EquityTheta):
    EtoD = E/D
    x0 = [1,1]
    def fun1(x):
        return KMVfun(EtoD, r, T, EquityTheta, x)
    return fsolve(fun1,x0)

if __name__ == "__main__":
    r = 0.0225  #无风险利率
    T=1         #期限一般设为1
    SD = 1e8
    LD = 50000000
    DP = 0.5*LD + SD #长期负债加短期负债等于总辐照
    D = DP
    princeTheta = 0.2893
    EquityTheta = princeTheta #公司的股权价值波动率
    E = 141276427  #公司的股权价值

    #江西赣粤高速公路股份有限公司
    # r = 0.0296  #无风险利率
    # T=1        #期限一般设为1
    # D = 10564057854.415  #债务
    # EquityTheta = 0.8989/100 #公司的股权价值波动率
    # E = 6585847779.48  #公司的股权价值

    #河南神火煤电股份有限公司
    # r = 0.0296  #无风险利率
    # T=1       #期限一般设为1
    # D = 30344052876.54  #债务
    # EquityTheta = 1.5939/100 #公司的股权价值波动率%
    # E = 8077125000  #公司的股权价值

    res = KMVOptSearch(E,D,r,T,EquityTheta)

    print("公司资产价值：",res[0]*E,"元\n","公司资产值波动率：",res[1])#结果
    print("资产股权比：",res[0])

    res_back = KMVfun(E/D,r,T,EquityTheta,res)
    eps = 1e-3
    if res_back[0] > eps or res_back[1] > eps:
        print("答案错误！")
