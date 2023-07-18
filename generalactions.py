from tkinter import messagebox
import pandas as pd
from СapitalСonstruction import cap_construction
from Opex import opex
from functools import cache
from drilling import drilling
from Excel import push_excel
import Global_Var
from equipment import OHCC
from Revex import revex
from Data import Data


class ParseAndEnter():
    def __init__(self, template, data):
        self.templatePath = template
        self.dataPath = data

    def init_const(self, excel, dfs):
        Global_Var.index_Nd = excel.find_index("НД")
        Global_Var.index_Service_Gov = excel.find_index("Служба ГС")
        Global_Var.index_direction_to_action = excel.find_index("Направление деятельности")
        Global_Var.index_None = excel.find_index(None)

        data = Data()
        dfs_data = (dfs.columns[6]).split(' ')[1]
        Global_Var.cur_data = data.current_time(dfs_data)
        Global_Var.next_month = data.next_time(dfs_data)

        Global_Var.columns_reserve = excel.find_column("Запасы на " + str(Global_Var.next_month), 5, 5)
        Global_Var.columns_profit = excel.find_column("Приход " + str(Global_Var.cur_data)[3:], 5, 5)
        if len(Global_Var.columns_profit) != 0:
            Global_Var.columns_lost = excel.find_column("Расход " + str(Global_Var.cur_data)[3:], 5, 5)

        Global_Var.start_cap_con = excel.find_row_direction_cases("КС")
        Global_Var.start_revex = excel.find_row_direction_cases("REVEX")
        Global_Var.start_opex = excel.find_row_direction_cases("OPEX")
        Global_Var.start_drilling = excel.find_row_direction_cases("Бурение")
        Global_Var.start_equipment = excel.find_row_direction_cases("ОНСС")
        print("const init")

    @cache
    def automatic(self):
        try:
            dfs = pd.read_excel(io=self.dataPath,
                                engine='openpyxl',
                                sheet_name='Sheet1')
        except:
            messagebox.showerror("Ошибка", "Не верно выбраны файлы")
        excel = push_excel(self.templatePath)
        self.init_const(excel, dfs)

        Revex = revex(excel)
        Revex.automatic(dfs, self.templatePath)
        OHHC = OHCC(excel)
        OHHC.automatic(dfs, self.templatePath)
        CS = cap_construction(excel)
        CS.automatic(dfs, self.templatePath)
        dril = drilling(excel)
        dril.automatic(dfs, self.templatePath)
        Opex = opex(excel)
        Opex.automatic(dfs, self.templatePath)
