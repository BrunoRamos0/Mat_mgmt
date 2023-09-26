import pcpy
import csv
import locale
import pandas as pd
from datetime import datetime
from dateutil.relativedelta import relativedelta

locale.setlocale(locale.LC_TIME, locale.normalize('pt_BR.utf8'))

materials = pcpy.Materials()

mat_list = materials.mat_list

pos = pcpy.PurchaseOrders(materials=materials)

def csv_tolist(filepath, float_column, date_column=None):

    dateformat = '%d/%m/%Y'

    with open(filepath, 'r') as file:
        file_list = list(csv.reader(file, delimiter=';'))

    file_list.pop(0)
    file_list = pcpy.txt_to_float(file_list, float_column)

    if date_column != None:
        for line in file_list:
            line[date_column] = datetime.strptime(line[date_column], dateformat)
    return file_list

# pos = csv_tolist('data/POs.csv', 3, 2)

cons = csv_tolist('data/cons.csv', 3, 0)

end_date = datetime(year=2024, month=3, day=31)

po_coverage = [['material', 'po', 'po_date', 'po_qty', 'po_inv']]

for mat, item in mat_list.items():

    date = datetime.today()
    cod = mat
    inv = item.inv

    pos_cod = pcpy.search_list(pos, cod, 4)
    cons_cod = pcpy.search_list(cons, cod, 1)


    for po in pos_cod:
        po_date = po[5]
        po_qty = po[6] - po[7]
        po_inv = pcpy.inv_po(date, po_date, cons_cod, inv, end_date)
        inv = po_inv + po_qty
        date = po_date
        po_coverage.append([cod, f'{po[0]}/{po[1]}', po_date, po_qty, po_inv])


df_po_coverage = pd.DataFrame(po_coverage[1:], columns=po_coverage[0])
df_po_coverage.to_csv('data/po_coverage.csv', sep=';', decimal=',', encoding='utf-8')