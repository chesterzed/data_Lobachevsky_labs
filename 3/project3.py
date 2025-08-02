import numpy as np
import xlrd
import matplotlib
import matplotlib.pyplot as plt
from math import sqrt


matplotlib.use('TkAgg')
############################
######## first task ########
############################


def get_math_expectation(x_p):
    res = 0
    for i in x_p:
        if i != '':
            res += float(i)
    res /= len(x_p)
    return res


def get_dispersion(x_p):
    d = []
    for i in x_p:
        if i != '':
            d.append(i ** 2)
    e1 = get_math_expectation(d)
    e2 = get_math_expectation(x_p) ** 2
    return e1 - e2


def get_covariatoin(x_p1, x_p2):
    c = 0
    e1 = get_math_expectation(x_p1)
    e2 = get_math_expectation(x_p2)
    for i in range(len(x_p1)):
        if x_p1[i] != '' or x_p2[i] != '':
            c += (x_p1[i] - e1) * (x_p2[i] - e2)
    c /= len(x_p1)
    return c


def get_correlation(x_p1, x_p2):
    c = get_covariatoin(x_p1, x_p2)
    d1 = sqrt(get_dispersion(x_p1))
    d2 = sqrt(get_dispersion(x_p2))
    r = c / (d1 * d2)
    return r


def cov_matrix(data_lists):
    li = []
    for i in range(len(data_lists)):
        li.append([])
        for j in range(len(data_lists)):
            c = get_covariatoin(data_lists[i], data_lists[j])
            li[i].append(c)
    return li


def cor_matrix(data_lists):
    li = []
    for i in range(len(data_lists)):
        li.append([])
        for j in range(len(data_lists)):
            c = get_correlation(data_lists[i], data_lists[j])
            li[i].append(c)
    return li


workbook = xlrd.open_workbook("02.xls")
sh = workbook.sheet_by_index(0)

# 3  - severity
# 18 - visibility
# 20 - wind_speed
# 16 - humidity
# 14 - temperature
num_ID = [3, 18, 20, 16, 14]
tf_ID = [23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34]

# get full table
names = []
values = []
for i in range(sh.row_len(0)):
    names.append(sh.cell_value(0, i))
    values.append(sh.col_values(i, start_rowx=1))

# get full filled number columns to  'num_vals = []'
num_vals = []
tf_vals = []

convert = {'True': 1, 'False': 0}

for i in range(len(num_ID)):
    num_vals.append([])

for i in range(len(tf_ID)):
    tf_vals.append([])

for i in range(len(values[0])):
    g = True
    for j in num_ID:
        if values[j][i] == '':
            g = False
            break
    for j in tf_ID:
        if values[j][i] == '':
            g = False
            break
    if g:
        for j in range(len(num_vals)):
            num_vals[j].append(float(values[num_ID[j]][i]))
        for j in range(len(tf_vals)):
            tf_vals[j].append(convert[values[tf_ID[j]][i]])


############################
####### second task ########
############################

# tests
print('-----mean-----')
print(str(np.mean(num_vals[0])))
print(get_math_expectation(num_vals[0]))

print('-----dispersion-----')
print(np.std(num_vals[0]) ** 2)
print(get_dispersion(num_vals[0]))

print('-----covariatoin-----')
print(np.var([num_vals[0], num_vals[0]]))
print(np.var(num_vals[0]))
print(get_covariatoin(num_vals[0], num_vals[0]))

print('-----cov_matrix-----')
m = np.cov(num_vals)
n = cov_matrix(num_vals)
print(str(m))
[print(i) for i in n]

print('-----correlation-----')
print(np.corrcoef(num_vals[0], num_vals[2])[0][1])
print(get_correlation(num_vals[0], num_vals[2]))

print('-----cor_matrix-----')
print(np.corrcoef(num_vals))
c = cor_matrix(num_vals)
[print(i) for i in c]

############################
######## third task ########
############################

fig, axes = plt.subplots(5, 5, figsize=(14, 9))

for i in range(5):
    for j in range(5):
        axes[j][i].scatter(num_vals[i], num_vals[j])
        axes[j][i].set_xlabel(names[num_ID[i]])
        if i == 0:
            axes[j][i].set_ylabel(names[num_ID[j]])

plt.show()

############################
####### fourth task ########
############################


# отметки о наличии вблизи места аварии
# лежачего полицейского, 
# перекрестка, 
# знака «Уступи дорогу»,
# транспортной развязки,

# знака «нет выхода», 
# железнодорожных путей,
# кругового движения, 
# остановки общественного транспорта (автобусов, поездов и т.п.), знака «стоп», 
# знаков или других мер успокоения движения, 

# светофоров, 
# поворотной петли

fig, axes = plt.subplots(3, 4, figsize=(14, 8))

for i in range(12):
    axes[i % 3][i % 4].scatter(tf_vals[i], num_vals[0])
    axes[i % 3][i % 4].set_xlabel(names[tf_ID[i]])
    if i == 0:
        axes[j % 3][i % 4].set_ylabel(names[num_ID[0]])

plt.show()
