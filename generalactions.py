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
from SZ_GO_etc import Sz_Go_etc


class ParseAndEnter():
    def __init__(self, template, data):
        self.templatePath = template
        self.dataPath = data

    @staticmethod
    def find_time_data(dfs_data):
        return Data().current_time(dfs_data), Data().next_time(dfs_data)

    def init_const(self, excel, dfs):
        Global_Var.index_Nd = excel.find_index("НД")
        Global_Var.index_Service_Gov = excel.find_index("Служба ГС")
        Global_Var.index_direction_to_action = excel.find_index("Направление деятельности")
        Global_Var.index_None = excel.find_index(None)
        dfs_data = (dfs.columns[6]).split(' ')[1]
        Global_Var.cur_data, Global_Var.next_month = self.find_time_data(dfs_data)
        parse_time = Global_Var.cur_data.split(' ')
        Global_Var.OP_column = excel.find_column("ОП " + parse_time[1][:-1] + "ь" + " " + parse_time[2] + 'г.', 5, 5)
        Global_Var.columns_reserve = excel.find_column("Запасы на " + str(Global_Var.next_month), 5, 5)
        Global_Var.columns_profit = excel.find_column("Приход " + str(Global_Var.cur_data)[3:], 5, 5)
        if len(Global_Var.columns_profit) != 0:
            Global_Var.columns_lost = excel.find_column("Расход " + str(Global_Var.cur_data)[3:], 5, 5)
        Global_Var.start_cap_con = excel.find_row_direction_cases("КС")
        Global_Var.start_revex = excel.find_row_direction_cases("REVEX")
        Global_Var.start_opex = excel.find_row_direction_cases("OPEX")
        Global_Var.start_drilling = excel.find_row_direction_cases("Бурение")
        Global_Var.start_equipment = excel.find_row_direction_cases("ОНСС")
        Global_Var.start_etc = excel.find_row_direction_cases("запасы ГО, СО, СИЗ")
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
        mass_objs = [cap_construction(excel), OHCC(excel), opex(excel), Sz_Go_etc(excel), revex(excel), OHCC(excel),
                     drilling(excel)]
        for el in mass_objs:
            el.automatic(dfs, self.templatePath)
        if len(Global_Var.mistakes) != 0:
            error_message = "\n".join(str(x) for x in Global_Var.mistakes)
            error_message = "Не удалось заполнить следующие записи:\n " + error_message
            messagebox.showerror("Ошибка", error_message)
