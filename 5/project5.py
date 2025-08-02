import xlrd
import numpy as np
import matplotlib.pyplot as plt

from sklearn.cluster import KMeans, AgglomerativeClustering
import scipy.cluster.hierarchy as sch


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


def select_cols_in_rows_data(rows, selected_cols):
    return [[rows[i][j] for j in selected_cols] for i in range(len(rows))]


####################
### prepare data ###
####################

workbook = xlrd.open_workbook("5__2/02_autoAcidents.xls")
sh = workbook.sheet_by_index(0)

# for convertation
boolean = {'True': 1, 'False': 0, 'Day': 1, 'Night': 0, 'R': 1, 'L': 0}
types = [str, str, float, int, str, str, float, str, str, str, str, str, str, str, float, float, float, float, float, str, float, float, str, str, str, str, str, str, str, str, str, str, str, str, str, str, str]
num_cols = [2, 3, 6, 9, 14, 15, 16, 17, 18, 20, 21, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36]

# for shop
col_vals = []
row_vals = []
col_names = []

#import
col_names = sh.row_values(0)
for i in range(1, len(sh.col_values(0, start_rowx=1)) + 1):#10):#
    row = sh.row_values(i)
    if not '' in row:
        row_vals.append(row)

#to data types
row_len = len(row_vals[0])
for j in range(len(row_vals)):
    row_vals[j] = [types[i](row_vals[j][i]) for i in range(row_len)]
    for i in range(row_len):
        if row_vals[j][i] in boolean.keys():
            row_vals[j][i] = boolean[row_vals[j][i]]

#convert
col_vals = create_list_of_lists(row_len)
for i in range(len(row_vals)):
    for j in range(row_len):
        col_vals[j].append(row_vals[i][j])

print('data is ready')

####################
### errors count ###
####################

num_types = [int, float]
errors = []
num_clust = np.arange(1, 20)

num_r = select_cols_in_rows_data(row_vals, num_cols)
for i in np.arange(1, 20):
    kmeans = KMeans(n_clusters=i).fit(num_r)
    errors.append(kmeans.inertia_)


plt.plot(np.arange(1, 20), errors)
plt.title("суммарная ошибка кластеризации методом К средних")
plt.xlabel("Число кластеров")
plt.show()

####################
### first task a ###
####################

####################
### print graphs ###
####################

common_params = {
    "n_init": "auto",
    "random_state": 4,
}

task_a_cols = [3, 14, 15, 16, 17, 18, 20, 21]

j = 0
lnc = len(task_a_cols)
fig, axes = plt.subplots(nrows=lnc, ncols=lnc, figsize=(12, 12))

for i in range(lnc ** 2):
    if i % lnc == 0:
        if i != 0:
            j += 1
        axes[j][i % lnc].set_ylabel(col_names[task_a_cols[j]])
    axes[j][i % lnc].scatter(col_vals[task_a_cols[i % lnc]], col_vals[task_a_cols[j]])
    axes[j][i % lnc].set_xlabel(col_names[task_a_cols[i % lnc]])

plt.show()

########################################

j = 0
fig, axes = plt.subplots(nrows=lnc, ncols=lnc, figsize=(12, 12))

for i in range(lnc ** 2):
    if i % lnc == 0:
        if i != 0:
            j += 1
        axes[j][i % lnc].set_ylabel(col_names[task_a_cols[j]])
    y_pred = KMeans(n_clusters=5, **common_params).fit_predict(select_cols_in_rows_data(row_vals, [task_a_cols[i % lnc], task_a_cols[j]]))
    axes[j][i % lnc].scatter(col_vals[task_a_cols[i % lnc]], col_vals[task_a_cols[j]], c=y_pred)
    axes[j][i % lnc].set_xlabel(col_names[task_a_cols[i % lnc]])

plt.show()

####################
### first task b ###
####################

####################
### print graphs ###
####################


task_b_cols = [3, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34]


j = 0
lnc = len(task_b_cols)
fig, axes = plt.subplots(nrows=lnc, ncols=lnc, figsize=(12, 12))

for i in range(lnc ** 2):
    if i % lnc == 0:
        if i != 0:
            j += 1
        axes[j][i % lnc].set_ylabel(col_names[task_b_cols[j]])
    axes[j][i % lnc].scatter(col_vals[task_b_cols[i % lnc]], col_vals[task_b_cols[j]])
    axes[j][i % lnc].set_xlabel(col_names[task_b_cols[i % lnc]])

plt.show()

########################################

j = 0
fig, axes = plt.subplots(nrows=lnc, ncols=lnc, figsize=(12, 12))

for i in range(lnc ** 2):
    if i % lnc == 0:
        if i != 0:
            j += 1
        axes[j][i % lnc].set_ylabel(col_names[task_b_cols[j]])
    y_pred = KMeans(n_clusters=5, **common_params).fit_predict(select_cols_in_rows_data(row_vals, [task_b_cols[i % lnc], task_b_cols[j]]))
    axes[j][i % lnc].scatter(col_vals[task_b_cols[i % lnc]], col_vals[task_b_cols[j]], c=y_pred)
    axes[j][i % lnc].set_xlabel(col_names[task_b_cols[i % lnc]])

plt.show()

###################
### second task ###
###################

###################
#### functions ####
###################

def plot_dendrogram(model, **kwargs):
    counts = np.zeros(model.children_.shape[0])
    n_samples = len(model.labels_)
    for i, merge in enumerate(model.children_):
        current_count = 0
        for child_idx in merge:
            if child_idx < n_samples:
                current_count += 1
            else:
                current_count += counts[child_idx - n_samples]
        counts[i] = current_count

    linkage_matrix = np.column_stack(
        [model.children_, model.distances_, counts]
    ).astype(float)

    sch.dendrogram(linkage_matrix, **kwargs)

markers = ['.','v', '^', '<', '>', '1', '2', '3', '4', '8', 's', 'p', 'P', '*', 'h', 'H', '+', 'x', 'X', 'D', 'd', '|', '_']
colors = ['b', 'g', 'r', 'c', 'm', 'y', 'k', 'w']
def plotCompareClusters(n_clusters, labels1, labels2, p1, p2):
    fig, axes = plt.subplots(1, 2, figsize=(12, 10))
    for i in range(len(row_vals)):
        axes[0].plot(row_vals[i][p1], row_vals[i][p2], markers[19] ,c=colors[labels1[i]], markersize = 3)
        axes[1].plot(row_vals[i][p1], row_vals[i][p2], markers[19] ,c=colors[labels2[i]], markersize = 3)
    axes[0].set_title("K срдених\n" + str(n_clusters) + " кластера")
    axes[1].set_title("Иерархическая кластеризация\n" + str(n_clusters) + " кластера")
    axes[0].set_xlabel(col_names[p1])
    axes[1].set_xlabel(col_names[p1])
    axes[0].set_ylabel(col_names[p2])
    axes[1].set_ylabel(col_names[p2])
    plt.show()

##################
###### code ######
##################

#####################
### second task a ###
#####################
n = 5

hierarch = AgglomerativeClustering(n_clusters=n, metric='euclidean', linkage='ward', compute_distances=True).fit(select_cols_in_rows_data(row_vals, task_a_cols))
plt.title("Hierarchical Clustering Dendrogram")
plot_dendrogram(hierarch, truncate_mode="level", p=8)
plt.xlabel("Number of points in node (or index of point if no parenthesis).")
plt.show()


clust_amount = 2

for i in np.arange(0, len(task_a_cols)):
    for j in np.arange(i + 1, len(task_a_cols)):
        kmeans = KMeans(n_clusters=clust_amount).fit(select_cols_in_rows_data(row_vals, [task_a_cols[i], task_a_cols[j]]))
        hierarch = AgglomerativeClustering(n_clusters=clust_amount, metric='euclidean', linkage='ward').fit(select_cols_in_rows_data(row_vals, [task_a_cols[i], task_a_cols[j]]))
        print(kmeans.labels_, hierarch.labels_)
        plotCompareClusters(i, kmeans.labels_, hierarch.labels_, task_a_cols[i], task_a_cols[j])


#####################
### second task b ###
#####################

hierarch = AgglomerativeClustering(n_clusters=n, metric='euclidean', linkage='ward', compute_distances=True).fit(select_cols_in_rows_data(row_vals, task_b_cols))
plt.title("Hierarchical Clustering Dendrogram")
plot_dendrogram(hierarch, truncate_mode="level", p=8)
plt.xlabel("Number of points in node (or index of point if no parenthesis).")
plt.show()

clust_amount = 3

for i in np.arange(0, len(task_b_cols)):
    for j in np.arange(i + 1, len(task_b_cols)):
        kmeans = KMeans(n_clusters=clust_amount).fit(select_cols_in_rows_data(row_vals, [task_b_cols[i], task_b_cols[j]]))
        hierarch = AgglomerativeClustering(n_clusters=clust_amount, metric='euclidean', linkage='ward').fit(select_cols_in_rows_data(row_vals, [task_b_cols[i], task_b_cols[j]]))
        print(kmeans.labels_, hierarch.labels_)
        plotCompareClusters(i, kmeans.labels_, hierarch.labels_, task_b_cols[i], task_b_cols[j])


