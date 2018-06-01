from scipy.optimize import minimize
import numpy as np

def minimizer_func(beta, fit_func, x_vals, q, observations):
    return q*np.sum(np.abs(np.where(observations>fit_func(beta, x_vals),observations - fit_func(beta, x_vals),0))) + \
           (1-q)*np.sum(np.abs(np.where(observations<fit_func(beta, x_vals),observations - fit_func(beta, x_vals),0)))

def quantile_regression(fit_func, x_vals, observations, beta_init, bounds=None, q_value = 0.5):
    return minimize(minimizer_func, beta_init, args=(fit_func, x_vals, q_value, observations), bounds=bounds)

#API -
#Object oriented design
#R1 score
#Other diagnostic
#Multi parameter analysis
#Add fit predict
import numpy as np

def func(beta, x):
    return beta[0]*np.power(x,2)

if __name__ == '__main__':
    import matplotlib.pyplot as plt

    x = np.array(range(15))
    y = 1.2*np.array(range(15))**1.23 + np.random.rand(15)

    a = quantile_regression(func, x, y, [1], q_value=0.1)
    b = quantile_regression(func, x, y, [1], q_value=0.9)
    print(a)
    print("====")
    print(b)
    plt.figure()
    plt.scatter(x,y)
    plt.plot(x,x*a.fun+a.jac[0])
    plt.plot(x,x*b.fun+b.jac[0])

    plt.show()