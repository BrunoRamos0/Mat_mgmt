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

po_list = pos.POs

end_date = datetime(year=2024, month=3, day=31)
start_date = datetime.today()
po_list = pcpy.remove_outofdate(nested_list=po_list, date_column=5,start_date=start_date, last_date=end_date)
po_list = sorted(po_list, key=lambda x: x[5])

cons = pcpy.csv_tolist('data/cons.csv', [3], 0)

po_coverage = [['material', 'po', 'po_date', 'po_qty', 'po_inv', 'po_cov']]

for mat, item in mat_list.items():

    date = datetime.today()
    cod = mat
    inv = item.inv

    pos_cod = pcpy.search_list(po_list, cod, 4)
    cons_cod = pcpy.search_list(cons, cod, 1)
    cons_cod = sorted(cons_cod, key=lambda x: x[0])

    if not pos_cod: continue

    for po in pos_cod:
        po_date = po[5]
        po_qty = po[6] - po[7]
        po_inv = pcpy.inv_po(date, po_date, cons_cod, inv, end_date)
        po_cov = pcpy.inv_po_coverage(po_inv, po_date, cons_cod, end_date)
        inv = po_inv + po_qty
        date = po_date
        po_coverage.append([cod, f'{po[1]}/{po[2]}', po_date, po_qty, po_inv, po_cov*30])


df_po_coverage = pd.DataFrame(po_coverage[1:], columns=po_coverage[0])
df_po_coverage.to_csv('data/po_coverage.csv', sep=';', decimal=',', encoding='utf-8', index=False)