import csv
import xlwt
from datetime import datetime as dt

# TODO Сделать из этого безобразия класс

reference = {}

with open("alias.csv") as csv_file:
    reader = csv.reader(csv_file, delimiter=';')
    for row in reader:
        reference[row[0]] = {
            'addresses': row[1].split(','),
            'domain_name': row[5],
            'date_reg': row[6],
            'date_edit': row[7],
            'date_valid': row[8],
            'not_box': int(row[10])
        }

for alias in reference.keys():
    reference[alias]['links'] = {}
    for address in reference[alias]['addresses']:
        if alias != address:
            link = reference.get(address)
            if link:
                reference[alias]['links'][address] = link

# print(reference['staff-slm@diasoft-service.ru']['not_box'])
# for val in reference.items():
#     print(val)

sorting_excel = [
    'service_managers@diasoft-service.ru',
    'support@diasoft-service.ru',
    'info@diasoft-service.ru',
    'service@diasoft-service.ru',
    'office@diasoft-service.ru',
    'paid@diasoft-service.ru',
    'index@diasoft-service.ru',
    'spares@diasoft-service.ru',
    'supsoft@diasoft-service.ru',
    'admin@diasoft-service.ru',
    'slm@diasoft-service.ru',
    'flm@diasoft-service.ru',
    'coor@diasoft-service.ru',
    'coor-slm@diasoft-service.ru',
    'coor-flm@diasoft-service.ru',
    'staff-slm@diasoft-service.ru',
    'staff-slm-krsk@diasoft-service.ru',
    'staff-flm@diasoft-service.ru',
    'staff-flm-krsk@diasoft-service.ru',
    'staff-krsk@diasoft-service.ru',
    'drm@diasoft-service.ru',
]
stack = reference.copy()

font_title = xlwt.Font()
font_title.bold = True
stype_title = xlwt.XFStyle()
stype_title.font = font_title


def cell_write(ws, r, c, label='', *args, **kwargs):
    ws.write(r, c, label, *args, **kwargs)
    width = int((1 + len(str(label))) * 240)
    if width > ws.col(c).width:
        ws.col(c).width = width

wb = xlwt.Workbook()
ws = wb.add_sheet('Рассылки')
for col, address in enumerate(sorting_excel):
    alias = stack.pop(address)
    # print(col, address, alias["addresses"])
    cell_write(ws, 0, col, address, stype_title)
    for row, link in enumerate(alias['links'], 1):
        cell_write(ws, row, col, link)
other_alias_list = filter(
    lambda alias: alias[1]['domain_name'] == 'diasoft-service.ru' and len(alias[1]['links']) > 0,
    stack.items()
)
other_alias_list = sorted(other_alias_list, key=lambda alias: len(alias[1]['links']), reverse=True)

for col, alias in enumerate(other_alias_list, col + 1):

    cell_write(ws, 0, col, alias[0], stype_title)
    for row, link in enumerate(alias[1]['links'], 1):
        cell_write(ws, row, col, link)

out_file_name = "Рсааылки_{}.xls".format(dt.now().strftime("%Y.%m.%d"))
wb.save(out_file_name)
print('---')
print()


# for key, val in other_alias_list:
#     print("{:<40}|{}".format(key, len(val['links'])))


# out_file_name = "Рсааылки_{}.csv".format(dt.now().strftime("%Y.%m.%d"))
# with open(out_file_name, 'w') as csv_file:
#     writer = csv.writer(csv_file)
#     writer.writerows(not_box_list)

# not_box_list = filter(
#     lambda item: item[1]['not_box'] == 1 and item[1]['domain_name'] == 'diasoft-service.ru',
#     reference.items()
# )

# for key, val in not_box_list:
#     print("'", key, "',", sep='')

# not_box_list_sort = sorted(not_box_list, key=lambda item: len(item[1]['links']), reverse=True)

# for key, val in not_box_list_sort:
#     print("{:<40}|{}".format(key, len(val['links'])))

# with open("alias_17.08.2016.csv") as f:
#     lines = f.readlines()
#
# for line in lines:
#     fields = line.split(";")
