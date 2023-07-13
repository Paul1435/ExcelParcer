from tkinter import messagebox
import pandas as pd
from СapitalСonstruction import cap_construction
from functools import cache


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
        CS = cap_construction()
        CS.automatic(dfs, self.templatePath)
