
import win32com.client
import os
from datetime import datetime
from dateutil.relativedelta import relativedelta

class SAP_Update():
    
    def __init__(self):

        self.SapGuiAuto = win32com.client.GetObject('SAPGUI')
        if not type(self.SapGuiAuto) == win32com.client.CDispatch:
            return

        self.application = self.SapGuiAuto.GetScriptingEngine
        if not type(self.application) == win32com.client.CDispatch:
            self.SapGuiAuto = None
            return
        
        self.connection = self.application.Children(0)
        if not type(self.connection) == win32com.client.CDispatch:
            self.application = None
            self.SapGuiAuto = None
            return

        self.session = self.connection.Children(0)
        if not type(self.session) == win32com.client.CDispatch:
            self.connection = None
            self.application = None
            self.SapGuiAuto = None
            return

    def Get_YP22(self, Comp, filepath=None):

        if filepath == None:
            filepath = os.getcwd() + '/data/'

        self.session.findById("wnd[0]/tbar[0]/okcd").Text = "/NYP22"
        self.session.findById("wnd[0]").sendVKey(0)
        self.session.findById("wnd[0]/usr/ctxtP_WERKS").Text = "CED"

        Comp.to_clipboard(excel=True, index=False)

        self.session.findById("wnd[0]/usr/btn%_S_MATNR_%_APP_%-VALU_PUSH").press()
        self.session.findById("wnd[1]").sendVKey(24)
        self.session.findById("wnd[1]").sendVKey(8)

        if False:
            self.session.findById("wnd[0]/usr/btn%_S_MATNRA_%_APP_%-VALU_PUSH").press()
            self.session.findById("wnd[1]").sendVKey(24)
            self.session.findById("wnd[1]").sendVKey(8)
            
        self.session.findById("wnd[0]/usr/chkP_PERDA").Selected = True
        self.session.findById("wnd[0]/usr/chkP_MEHRS").Selected = True
        self.session.findById("wnd[0]/usr/chkP_OTWRK").Selected = False
        self.session.findById("wnd[0]/usr/chkP_ESTQ").Selected = False
        self.session.findById("wnd[0]").sendVKey (8)
        self.session.findById("wnd[0]/tbar[1]/btn[45]").press()
        self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
        self.session.findById("wnd[1]/usr/ctxtDY_PATH").Text = filepath
        self.session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "DATA.TXT"
        self.session.findById("wnd[1]/tbar[0]/btn[11]").press()

    def Get_Inv(self, Comp, filepath=None):

        if filepath == None:
            filepath = os.getcwd() + '/data/'

        self.session.findById("wnd[0]/tbar[0]/okcd").Text = "/nmb52"
        self.session.findById("wnd[0]").sendVKey(0)
        self.session.findById("wnd[0]").sendVKey(17)
        self.session.findById("wnd[1]/usr/txtV-LOW").Text = "sext"
        self.session.findById("wnd[1]").sendVKey(8)

        Comp.to_clipboard(excel=True, index=False)

        self.session.findById("wnd[0]/usr/btn%_MATNR_%_APP_%-VALU_PUSH").press()
        self.session.findById("wnd[1]").sendVKey(24)
        self.session.findById("wnd[0]").sendVKey(8)
        self.session.findById("wnd[0]/usr/ctxtWERKS-LOW").Text = "ced"
        self.session.findById("wnd[0]").sendVKey(0)
        self.session.findById("wnd[0]").sendVKey(8)
        self.session.findById("wnd[0]/tbar[1]/btn[45]").press()
        self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
        self.session.findById("wnd[1]/usr/ctxtDY_PATH").Text = filepath
        self.session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "data.txt"
        self.session.findById("wnd[1]/tbar[0]/btn[11]").press()

        filepath = filepath + 'data.txt'
        return filepath

    def get_PEDPEND(self, Comp, filepath=None):

        if filepath == None:
            filepath = os.getcwd() + '/data/'

        self.session.findById("wnd[0]/tbar[0]/okcd").Text = "/nsq00"
        self.session.findById("wnd[0]").sendVKey(0)
        self.session.findById("wnd[0]/tbar[1]/btn[19]").press()
        self.session.findById("wnd[1]/usr/cntlGRID1/shellcont/shell").SetCurrentCell(6, "DBBGTEXT")
        self.session.findById("wnd[1]/usr/cntlGRID1/shellcont/shell").SelectedRows = "6"
        self.session.findById("wnd[1]/usr/cntlGRID1/shellcont/shell").DoubleClickCurrentCell()
        self.session.findById("wnd[0]/usr/ctxtRS38R-QNUM").Text = "PED_PEND"
        self.session.findById("wnd[0]/usr/ctxtRS38R-QNUM").SetFocus()
        self.session.findById("wnd[0]/usr/ctxtRS38R-QNUM").caretPosition = 8
        self.session.findById("wnd[0]").sendVKey(8)
        self.session.findById("wnd[0]").sendVKey(17)
        self.session.findById("wnd[1]").sendVKey(8)
        
        Comp.to_clipboard(excel=True, index=False)

        self.session.findById("wnd[0]/usr/btn%_SP$00010_%_APP_%-VALU_PUSH").press()
        self.session.findById("wnd[1]").sendVKey(16)
        self.session.findById("wnd[1]").sendVKey(24)
        self.session.findById("wnd[1]").sendVKey(8)
        self.session.findById("wnd[0]").sendVKey(8)
        self.session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").PressToolbarContextButton("&MB_EXPORT")
        self.session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").SelectContextMenuItem("&PC")
        self.session.findById("wnd[1]").sendVKey(0)
        self.session.findById("wnd[1]/usr/ctxtDY_PATH").Text = filepath
        self.session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "data.txt"
        self.session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 8
        self.session.findById("wnd[1]/tbar[0]/btn[11]").press()

        filepath = filepath + 'data.txt'
        return filepath

    def get_MARC(self, Comp, filepath=None):

        if filepath == None:
            filepath = os.getcwd() + '/data/'

        Comp.to_clipboard(excel=True, index=False)

        self.session.findById("wnd[0]/tbar[0]/okcd").text = "/NSE16N"
        self.session.findById("wnd[0]").sendVKey(0)
        self.session.findById("wnd[0]/usr/ctxtGD-TAB").text = "MARC"
        self.session.findById("wnd[0]").sendVKey(0)
        self.session.findById("wnd[0]/usr/txtGD-MAX_LINES").text = ""
        self.session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/btnPUSH[4,1]").press()
        self.session.findById("wnd[1]").sendVKey(24)
        self.session.findById("wnd[1]").sendVKey(8)
        self.session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/ctxtGS_SELFIELDS-LOW[2,2]").text = "CED"
        self.session.findById("wnd[0]").sendVKey(8)
        self.session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").pressToolbarContextButton("&MB_EXPORT")
        self.session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").selectContextMenuItem("&PC")
        self.session.findById("wnd[1]").sendVKey(0)
        self.session.findById("wnd[1]/usr/ctxtDY_PATH").Text = filepath
        self.session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "DATA.TXT"
        self.session.findById("wnd[1]/tbar[0]/btn[11]").press()
		
    def get_j3(self, year = datetime.now().year, month = datetime.now().month, filepath=None):

        if filepath == None:
            filepath = os.getcwd() + '/data/'

        endday = SAP_Update.end_month(year=year, month=month)
        startdate = f'01.{month}.{year}'        
        enddate = f'{endday}.{month}.{year}'
        print(startdate, enddate)
        print('Carregando dados do SAP...')
        self.session.findById("wnd[0]/tbar[0]/okcd").text = "/nyp10"
        self.session.findById("wnd[0]").sendVKey(0)
        self.session.findById("wnd[0]/usr/chkP_ZEROS").selected = True
        self.session.findById("wnd[0]/usr/ctxtS_PRGRP-LOW").text = "acabamento"
        self.session.findById("wnd[0]/usr/ctxtP_WERKS").text = "dex"
        self.session.findById("wnd[0]/usr/ctxtP_INIDAT").text = startdate
        self.session.findById("wnd[0]/usr/ctxtP_FIMDAT").text = enddate
        self.session.findById("wnd[0]/usr/ctxtP_WERKSP").text = "dex"
        self.session.findById("wnd[0]/usr/txtP_VERSB").text = "j3"
        self.session.findById("wnd[0]").sendVKey(0)
        self.session.findById("wnd[0]").sendVKey(8)
        self.session.findById("wnd[0]").sendVKey(18)
        self.session.findById("wnd[0]").sendVKey(20)
        self.session.findById("wnd[1]").sendVKey(0)
        self.session.findById("wnd[1]/usr/ctxtDY_PATH").text = filepath
        self.session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "j3.txt"
        self.session.findById("wnd[1]/tbar[0]/btn[11]").press()

            
    @staticmethod
    def end_month(year, month):
        dateSart = datetime(year, month, 1)
        dateEnd = dateSart + relativedelta(day=31)
        return dateEnd.day


class SAP_Parse():

    @staticmethod
    def parse_YP22(filepath, writepath=r'data/parsed_YP22.txt'):
        with open(filepath, "r") as f_input, open(writepath, "w") as f_output:

            f_output.write('Cen.|LTAlt.|SKU|UMB|Qtd.básica|Componente|Texto breve objeto|Qtd.(UMB)|UM\n')

            for line in filter(
                lambda x: len(x) > 2
                and x[0] == "|"
                and x[1] != "-"
                and not x[6].isalpha(),
                f_input
            ):
                ## Split on | delimeter
                line_contents = [x.strip() for x in line.split("|")]

                ## Replace '.' thousand separator with ''
                line_contents = line_contents[:5] + [line_contents[5].replace('.', '')] + \
                                line_contents[6:-3] + [word.replace('.', '') for word in line_contents[-3:]]

                f_output.write("|".join(line_contents[1:-1]) + "\n")

            return writepath 

    @staticmethod
    def parse_Inv(filepath, writepath=r'data/parsed_Inv.csv'):
        with open(filepath, "r") as f_input, open(writepath, "w") as f_output:

            for line in filter(
                lambda x: len(x) > 2
                and x[0] == "|"
                and x[1] != "-",
                f_input
            ):
                # Split on | delimeter
                line_contents = [x.strip() for x in line.split("|")]

                # Replace '.' thousand separator with ''
                line_contents = line_contents[:5] + [word.replace('.', '') for word in line_contents[5:]]

                f_output.write(";".join(line_contents[1:-1]) + "\n")
            
            return writepath

    @staticmethod
    def parse_PEDPEND(filepath, writepath=r'data/parsed_PEDPEND.csv'):
        with open(filepath, "r") as f_input, open(writepath, "w") as f_output:

            for line in filter(
                lambda x: len(x) > 2
                and x[0] == "|"
                and x[1] != "-"
                and x[3] != " ",
                f_input
            ):
                # Split on | delimeter
                line_contents = [x.strip() for x in line.split("|")]

                # Replace '.' thousand separator with ''
                line_contents = line_contents[:15] + [word.replace('.', '') for word in line_contents[15:-17]] + line_contents[-17:]
                line_contents = line_contents[1:-1]

                if not line_contents[0]:
                    continue

                # Keep only selected columns
                column_list = [1, 2, 3, 5, 9, 13, 15, 17, 18, 20]
                line_contents = [line_contents[i] for i in column_list]

                f_output.write(";".join(line_contents) + "\n")

            return writepath  

    @staticmethod
    def parse_j3(filepath, writepath=r'data/parsed_j3.txt'):
        with open(filepath, "r") as f_input, open(writepath, "w") as f_output:

            for line in filter(
                lambda x: len(x) > 2
                and x[0] == "|"
                and x[1] != "-"
                and x[4] == " "
                and x[5] != " ",
                f_input
            ):
                # Split on | delimeter
                line_contents = [x.strip() for x in line.split("|")]

                f_output.write("|".join(line_contents[1:5]) + "\n")
            
            return writepath  

    @staticmethod
    def parse_Hist(filepath, writepath=r'data/parsed_Hist.txt'):
        with open(filepath, "r") as f_input, open(writepath, "w") as f_output:

            for line in filter(
                lambda x: len(x) > 2
                and x[0] == "|"
                and x[1] != "-",
                f_input
            ):
                # Split on | delimeter
                line_contents = [x.strip() for x in line.split("|")]

                # Replace '.' thousand separator with ''
                line_contents = line_contents[:15] + [word.replace('.', '') for word in line_contents[15:-17]] + line_contents[-17:]

                f_output.write("|".join(line_contents[1:-1]) + "\n")

            return writepath  

