import xlrd


def check(ignore_list, text):
    res = False
    for item in ignore_list:
        if item in text:
            res = True
            break
    return res


ignore = ['федеральный']

book1 = xlrd.open_workbook("01_01.xls")
book2 = xlrd.open_workbook("01_02.xls")
d1 = dict()

sh1 = book1.sheet_by_index(0)
sh2 = book2.sheet_by_index(0)

for i in range(4, 104):
    t = sh2.row_values(i, 0, 3)
    if not check(ignore, t[0]):
        d1[t[0]] = t[1:]

for i in range(4, 98):
    t = sh1.row_values(i, 0, 3)
    if not check(ignore, t[0]):
        if t[0] in d1.keys():
            d1[t[0]].insert(1, t[1])
            d1[t[0]].insert(2, t[2])
        else:
            d1[t[0]] = t[1:]

book1.release_resources()
book2.release_resources()

result = dict(sorted(d1.items(), key=lambda x: sum([0 if i == '' else i for i in x[1]]) / len(x)))

for k, v in result.items():
    print(f'{k:40}: {str(v):>35}')
