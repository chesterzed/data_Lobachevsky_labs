import xlrd
import matplotlib.pyplot as plt
import seaborn as sns
from collections import Counter


############################
def get_type(sh, i, j):
    cll = sh.cell_value(rowx=i, colx=j)
    try:
        if xlrd.xldate.xldate_as_datetime(sh.cell_value(rowx=i, colx=j), workbook.datemode).year > 2010:
            cll = xlrd.xldate.xldate_as_datetime(sh.cell_value(rowx=i, colx=j), workbook.datemode)
        else:
            raise TypeError
    except TypeError:
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
    # print(cll, type(cll))
    return type(cll)

############################

workbook = xlrd.open_workbook("02.xls")
sh = workbook.sheet_by_index(0)

############################
######## first task ########
############################

data_types = []
t = True
j = 1

while t:
    for i in range(sh.row_len(1)):
        # print(sh.cell_value(j, i))
        if len(data_types) <= i:
            data_types.append(get_type(sh, j, i))
        else:
            if data_types[i] == '':
                data_types[i] = get_type(sh, j, i)
    if not '' in data_types:
        t = False

for i in range(len(data_types)):
    print(f'{i})', data_types[i])
# print(data_types)

############################
####### second task ########
############################

vis_vals = sh.col_values(18, start_rowx=1)
severity_vals = sh.col_values(3, start_rowx=1)
# print(severity_vals)
full_vis = sorted([float(x) for x in vis_vals if x != ''])

vis1 = sorted([float(vis_vals[i]) for i in range(len(vis_vals)) if vis_vals[i] != '' and severity_vals[i] == '1'])
vis2 = sorted([float(vis_vals[i]) for i in range(len(vis_vals)) if vis_vals[i] != '' and severity_vals[i] == '2'])
vis3 = sorted([float(vis_vals[i]) for i in range(len(vis_vals)) if vis_vals[i] != '' and severity_vals[i] == '3'])
vis4 = sorted([float(vis_vals[i]) for i in range(len(vis_vals)) if vis_vals[i] != '' and severity_vals[i] == '4'])
sns.distplot(full_vis, hist=False, kde=True, label='all')
sns.distplot(vis1, hist=False, kde=True, label='severity 1')
sns.distplot(vis2, hist=False, kde=True, label='severity 2')
sns.distplot(vis3, hist=False, kde=True, label='severity 3')
sns.distplot(vis4, hist=False, kde=True, label='severity 4')
plt.legend()
plt.show()

sns.distplot(vis1, hist=True, kde=True, label='severity 1')
sns.distplot(vis2, hist=True, kde=True, label='severity 2')
sns.distplot(vis3, hist=True, kde=True, label='severity 3')
sns.distplot(vis4, hist=True, kde=True, label='severity 4')
plt.legend()
plt.show()


############################
######## third task ########
############################

temperature = sh.col_values(14, start_rowx=1)
variation_temperature_series = sorted([float(x) for x in temperature if x != ''])
stat_temperature_series = dict(Counter(variation_temperature_series))

maximum = list(stat_temperature_series.keys())[-1]
step = 5
group_stat_series = dict()
for i in range(1, int(maximum // step) + 3):
    group_stat_series[i * step - step/2] = 0

for k in stat_temperature_series.keys():
    s = (k // step) * step + step/2.
    if k % step != 0:
        group_stat_series[s] += stat_temperature_series[k]
    else:
        group_stat_series[s] += stat_temperature_series[k] / 2
        s -= step
        group_stat_series[s] += stat_temperature_series[k] / 2

print(stat_temperature_series)
print(group_stat_series)

sns.distplot(variation_temperature_series, hist=True, kde=True, label='temperature')
plt.legend()
plt.show()


############################
####### fourth task ########
############################

cities = sh.col_values(10, start_rowx=1)
distance = sh.col_values(6, start_rowx=1)

# get top five
dict_cities = dict(Counter(cities))
sc = sorted(dict_cities.items(), key=lambda x:x[1])
top5 = [sc[i][0] for i in range(-1, -6, -1)]
print(top5)

# get distances
d1 = sorted([float(distance[i]) for i in range(len(distance)) if distance[i] != '' and cities[i] == top5[0]])
d2 = sorted([float(distance[i]) for i in range(len(distance)) if distance[i] != '' and cities[i] == top5[1]])
d3 = sorted([float(distance[i]) for i in range(len(distance)) if distance[i] != '' and cities[i] == top5[2]])
d4 = sorted([float(distance[i]) for i in range(len(distance)) if distance[i] != '' and cities[i] == top5[3]])
d5 = sorted([float(distance[i]) for i in range(len(distance)) if distance[i] != '' and cities[i] == top5[4]])

sns.distplot(d1, hist=True, kde=True, label=top5[0])
sns.distplot(d2, hist=True, kde=True, label=top5[1])
sns.distplot(d3, hist=True, kde=True, label=top5[2])
sns.distplot(d4, hist=True, kde=True, label=top5[3])
sns.distplot(d5, hist=True, kde=True, label=top5[4])
plt.legend()
plt.show()

############################
######## last task #########
############################

f = dict()
for i in range(0, 6):
    f[i] = 0

for i in range(1, len(cities) + 1):
    if sh.cell_value(i, 10) == top5[0]:
        factors = 0
        v = sh.cell_value(i, 18)
        w = sh.cell_value(i, 20)
        s = sh.cell_value(i, 22)
        c = sh.cell_value(i, 24)
        j = sh.cell_value(i, 26)
        t = sh.cell_value(i, 32)
        if v != '' and float(v) < 2.:
            factors += 1
        if 2!= '' and float(2) > 10.:
            factors += 1
        if 'Snow' in s:
            factors += 1
        if c == 'True' or j == 'True':
            factors += 1
        if t == 'True':
            factors += 1
        f[factors] += 1
# print(f)

y, x = zip(*f.items())
plt.bar(x, y, width=10)
plt.show()