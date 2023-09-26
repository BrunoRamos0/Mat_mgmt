import glob
from openpyxl import load_workbook
import pandas as pd
import numpy as np
import sapy
from datetime import datetime
from dateutil.relativedelta import relativedelta


class Plans():


    def __init__(self, folder='Plans', plan_type='main'):
        self.folder = folder
        self.planlist = []
        self.plan = {}
        self.plan_type = plan_type
        Plans.get_plans(self, self.folder)


    def get_plans(self, folder):
        files = glob.glob(folder + '/0*.xlsx')
        for file in files:
            self.planlist.extend(Plans.read_excel(self, file))


    def read_excel(self, file_path):

        df_excel = pd.read_excel(file_path, sheet_name='Plan1')
        df_excel.columns = df_excel.columns[:3].tolist() + pd.to_datetime(df_excel.columns[3:], dayfirst=True).date.tolist()
        df_excel = pd.melt(df_excel.reset_index(), id_vars=['Material', 'Plano'], value_vars=pd.date_range('2023/04/01', '2024/03/01', freq='MS').date, var_name='Date', value_name='Qtd')
        
        plan_excel = df_excel.values.tolist()

        return plan_excel


class Materials():

    def __init__(self, file_path='data/Class MPs.xlsx'):
        self.file_path = file_path
        self.materials = Materials.get_materials(self, self.file_path)
        self.inventory = Materials.get_inv(list(self.materials.keys()), update=False)
        self.agg_inv = Materials.agg_inv(self.inventory)

        self.mat_list = Materials.create_material(self)


    def get_materials(self, file_path):
        materials_list = pd.read_excel(file_path, sheet_name='Plan1')
        materials_list = materials_list.values.tolist()
        materials = {x[0]:x[1:] for x in materials_list}

        return materials
    

    def create_material(self):
        mat_list = {}

        for key in self.materials:
            mat_key = key
            desc = self.materials[key][0]
            mat_class = self.materials[key][3]
            inv = self.agg_inv['Utilizlivre'].loc[self.agg_inv['Material'] == mat_key].max()
            if np.isnan(inv): inv=0
            mat_list[mat_key] = Material(code=mat_key, description=desc, plan_class=mat_class, inv=inv)
        
        return mat_list


    @staticmethod
    def get_inv(code, update=True):

        codes = pd.Series(data=code)

        if update:
            filepath = sapy.SAP_Update().Get_Inv(Comp=codes)
            filepath = sapy.SAP_Parse.parse_Inv(filepath=filepath)
        else:
            filepath = 'data/parsed_Inv.txt'

        return pd.read_csv(filepath, sep='|', thousands='.', decimal=',', encoding='ISO-8859-1')
    

    @staticmethod
    def agg_inv(inventory):
        agg_inv = inventory[['Material', 'Utilizlivre', 'Em CtrQld', 'Bloqueado']].groupby('Material').sum().reset_index()
        return agg_inv
    

    @staticmethod
    def get_MARC(codes=None, type='current'):

        if (type != 'current') and (codes != None):
            codes = pd.Series(data=codes)
            sapy.SAP_Update.get_MARC(Comp=codes)


    @staticmethod
    def get_hist(codes=None, filepath='data/hist.csv'):

        if codes != None:
            pass

        return pd.read_csv(filepath, sep=';', thousands='.', decimal=',', encoding='ISO-8859-1', parse_dates=[1])


class Material():
    
    def __init__(self, code, description, plan_class, inv=0, pol=30):
        self.code = code
        self.description = description
        self.plan_class = plan_class
        self.inv = inv
        self.pol = pol


class PurchaseOrders():

    def __init__(self, codes=None, materials=None):
        
        if materials == None:
            pass

        self.materials = materials
        
        codes=list(self.materials.mat_list.keys())
        if False:
            self.POs = PurchaseOrders.get_POs(self, codes=codes, update=True)

        self.POs = PurchaseOrders.get_POs(self)


    def get_POs(self, codes=None, filepath='data/parsed_PEDPEND.txt', update=False):

        if update:
            codes = pd.Series(data=codes)
            filepath = sapy.SAP_Update().get_PEDPEND(Comp=codes)
            filepath = sapy.SAP_Parse.parse_PEDPEND(filepath=filepath)

        return pd.read_csv(filepath, sep='|', thousands='.', decimal=',', encoding='ISO-8859-1')
    

def search_list(list, text, column):
    list_final = []
    for line in list:
        if line[column] == text:
            list_final.append(line)

    return list_final

def txt_to_float(nested_list, column):
    for line in nested_list:
        line[column] = float(line[column].replace('.', '').replace(',', '.'))
    return nested_list
    
def inv_po(date, po_date, cons_cod, inv, end_date):
    eom = datetime(date.year, date.month, 1) + relativedelta(months=1, days=-1)
    som = datetime(date.year, date.month, 1)
    next_date = min(po_date, eom)
    days = next_date - date
    days = days.total_seconds() / 60 / 60 / 24
    cons = float(search_list(cons_cod, som, 0)[0][3])/30
    inv -= cons * days
    if next_date == po_date:
        return inv
    elif next_date > end_date:
        return inv
    else:
        next_date = next_date + relativedelta(days=1)
        return inv_po(next_date, po_date, cons_cod, inv, end_date)