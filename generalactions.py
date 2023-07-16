from tkinter import messagebox
import pandas as pd
from СapitalСonstruction import cap_construction
from Opex import opex
from functools import cache
from drilling import drilling
from Push_Excel import push_excel
import Global_Var
from equipment import OHCC


class ParseAndEnter():
    def __init__(self, template, data):
        self.templatePath = template
        self.dataPath = data

    @cache
    def automatic(self):
        try:
            dfs = pd.read_excel(io=self.dataPath,
                                engine='openpyxl',
                                sheet_name='Sheet1')
        except:
            messagebox.showerror("Ошибка", "Не верно выбраны файлы")
        excel = push_excel(self.templatePath)
        Global_Var.index_Nd = excel.find_index("НД")
        Global_Var.index_Service_Gov = excel.find_index("Служба ГС")
        Global_Var.index_direction_to_action = excel.find_index("Направление деятельности")
        Global_Var.index_None = excel.find_index(None)
        OHHC = OHCC(excel)
        OHHC.automatic(dfs, self.templatePath)
        # CS = cap_construction(excel)
        # CS.automatic(dfs, self.templatePath)
        # dril = drilling(excel)
        # dril.automatic(dfs, self.templatePath)
        # Opex = opex(excel)
        # Opex.automatic(dfs, self.templatePath)
