import xlrd
import numpy as np
import matplotlib.pyplot as plt
from sklearn.model_selection import train_test_split
from sklearn.metrics import mean_squared_error, r2_score
from sklearn.linear_model import LinearRegression


def get_type(sh, i, j):
    cll = sh.cell_value(rowx=i, colx=j)
    cll = sh.cell_value(rowx=i, colx=j)
    if cll == 'True':
        cll = True
    elif cll == 'False':
        cll = False
    else:
        try:
            cll = int(sh.cell_value(rowx=i, colx=j))
        except ValueError:
            try:
                cll = float(sh.cell_value(rowx=i, colx=j))
            except ValueError:
                cll = str(cll)
    return cll


def create_list_of_lists(count):
    li = []
    for i in range(count):
        li.append([])
    return li

workbook = xlrd.open_workbook("06_HousePrice.xls")
sh = workbook.sheet_by_index(0)

############################
######## first task ########
############################

data_types = []
t = True
j = 33

while t:
    for i in range(sh.row_len(1)):
        if len(data_types) <= i:
            data_types.append(get_type(sh, j, i))
        else:
            if data_types[i] == '':
                data_types[i] = get_type(sh, j, i)
    j += 1
    if not '' in data_types:
        t = False

for i in range(len(data_types)):
    print(f'{i})', type(data_types[i]))


############################
####### second task ########
############################

col_vals = []
names = []

print(float('1.225e+006'))

for i in range(sh.row_len(0)):
    names.append(sh.cell_value(0, i))
    col_vals.append(list(map(float, sh.col_values(i, start_rowx=1))))

print('data is ready')

fig, axes = plt.subplots(3, 6, figsize=(22, 12))

j = 0
for i in range(18):
    if i % 6 == 0 and i != 0:
        j += 1
    axes[j][i % 6].scatter(col_vals[2], col_vals[i])
    axes[j][i % 6].set_xlabel(names[i])
    axes[j][i % 6].set_ylabel(names[2])

print('ready to print')
plt.show()

############################
######## third task ########
############################


# transform data
nrows = len(col_vals[0])
transformed_data = create_list_of_lists(nrows)
for col in col_vals:
    for i in range(nrows):
        transformed_data[i].append(col[i])

# LinearRegression
price_num = 2
param_num = 5

priceTr, priceVal = train_test_split(transformed_data, test_size=0.2, train_size=0.8)
X = np.reshape(col_vals[price_num], (-1, 1))

linR = LinearRegression().fit(X, col_vals[param_num])

a = linR.coef_
b = linR.intercept_
print("Построена зависимость: Y = " + str(a[0]) + " X = " + str(b))

Y_predict = linR.predict(np.reshape(col_vals[price_num], (-1, 1)))

r2_1 = r2_score(col_vals[param_num], Y_predict)

print("Коэффициент детерминации = " + str(r2_1))

mse_1 = mean_squared_error(col_vals[param_num], Y_predict)
print ("Средняя квадратическая ошибка (MSE) = " + str(mse_1))

rss_1 = ((col_vals[param_num] - Y_predict)**2).sum()
print("Остаточная сумма квадратов (RSS) = " + str(rss_1))


plt.scatter(col_vals[price_num], col_vals[param_num])
plt.plot(col_vals[price_num], Y_predict, color='green')
plt.title('Модель линейной регрессии')
plt.xlabel(names[param_num])
plt.ylabel(names[price_num])
plt.show()

# Multi Linear Regression


