import csv

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
            'not_box': row[10]
        }

for alias in reference.keys():
    reference[alias]['links'] = {}
    for adress in reference[alias]['adresses']:
        if (alias != adress):
            link = reference.get(adress)
            if link:
                reference[alias]['links'][adress] = link

for key, val in reference.items():
    print(key, val)

# with open("alias_17.08.2016.csv") as f:
#     lines = f.readlines()
#
# for line in lines:
#     fields = line.split(";")