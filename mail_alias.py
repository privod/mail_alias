import csv
import xlwt
from datetime import datetime as dt

reference = {}

with open("alias.csv") as csv_file:
    reader = csv.reader(csv_file, delimiter=';')
    for row in reader:
        reference[row[0]] = {
            'adresses': row[1].split(','),
            'domen_name': row[5],
            'date_reg': row[6],
            'date_edit': row[7],
            'date_valid': row[8],
            'not_box': int(row[10])
        }

for alias in reference.keys():
    reference[alias]['links'] = {}
    for adress in reference[alias]['adresses']:
        if (alias != adress):
            link = reference.get(adress)
            if link:
                reference[alias]['links'][adress] = link

# print(reference['staff-slm@diasoft-service.ru']['not_box'])
# for val in reference.items():
#     print(val)

not_box_list = filter(lambda item: item[1]['not_box'] == 1 and item[1]['domen_name'] == 'diasoft-service.ru', reference.items())

print(enumerate(not_box_list))

wb = xlwt.Workbook()
ws = wb.add_sheet('Рассылки')
for i, item in list(enumerate(not_box_list)):
    print(i, item)
    # ws.write(0, 0, key)

# out_file_name = "Рсааылки_{}.csv".format(dt.now().strftime("%Y.%m.%d"))
# with open(out_file_name, 'w') as csv_file:
#     writer = csv.writer(csv_file)
#     writer.writerows(not_box_list)

not_box_list_sort = sorted(not_box_list, key=lambda item: len(item[1]['links']), reverse=True)


for key, val in not_box_list_sort:
    print("{:<40}|{}".format(key, len(val['links'])))

# with open("alias_17.08.2016.csv") as f:
#     lines = f.readlines()
#
# for line in lines:
#     fields = line.split(";")