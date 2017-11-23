'''
Useful regression handles
'''


import statsmodels.api as sm


def ols(data, x, y, add_constant=True):
    X = data[x]
    if add_constant:
        X = sm.add_constant(X)
    y = data[y]
    model = sm.OLS(y, X).fit()
    return model