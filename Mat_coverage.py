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

with open('data/POs.csv', 'r') as file:
    pos = list(csv.reader(file, delimiter=';'))
    pos.pop(0)

pos = pcpy.txt_to_float(pos, 3)

with open('data/cons.csv', 'r') as file:
    cons = list(csv.reader(file, delimiter=';'))
    cons.pop(0)

cons = pcpy.txt_to_float(cons, 3)

dateformat = '%d/%m/%Y'

for line in cons:
    line[0] = datetime.strptime(line[0], dateformat)

end_date = datetime(year=2024, month=3, day=31)

def inv_po(date, po_date, cons, inv):
    eom = datetime(date.year, date.month, 1) + relativedelta(months=1, days=-1)
    som = datetime(date.year, date.month, 1)
    next_date = min(po_date, eom)
    days = next_date - date
    days = days.total_seconds() / 60 / 60 / 24
    cons = float(pcpy.search_list(cons_cod, som, 0)[0][3])/30
    inv -= cons * days
    if next_date == po_date:
        return inv
    elif next_date > end_date:
        return inv
    else:
        next_date = next_date + relativedelta(days=1)
        return inv_po(next_date, po_date, cons, inv)

po_coverage = [['material', 'po', 'po_date', 'po_qty', 'po_inv']]

for mat, item in mat_list.items():

    date = datetime.today()
    cod = mat
    inv = item.inv

    pos_cod = pcpy.search_list(pos, cod, 1)
    cons_cod = pcpy.search_list(cons, cod, 1)


    for po in pos_cod:
        po_date = datetime.strptime(po[2], dateformat)
        po_qty = po[3]
        po_inv = inv_po(date, po_date, cons_cod, inv)
        inv = po_inv + po_qty
        date = po_date
        po_coverage.append([cod, po[0], po_date, po_qty, po_inv])

df_po_coverage = pd.DataFrame(po_coverage[1:], columns=po_coverage[0])
df_po_coverage.to_csv('data/po_coverage.csv', decimal=',', encoding='utf-8')