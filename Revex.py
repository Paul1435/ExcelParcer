from functools import cache
from Pivot_Table import create_pivot_table
from tkinter import messagebox
import Global_Var


class revex():
    def __init__(self, excel):
        self.pivot_table = None
        self.group_direction = ["ОД_вспомогательные", "Основная деятельность"]
        self.direction_do = [60]
        self.excel = excel
        print("revex init")

    @cache
    def find_row(self, sheet, Nd_requirements, GS_requirements, dta_requirements, fact_requirements, row_begin):
        for row in range(row_begin, sheet.max_row + 1):
            if Nd_requirements == sheet[row][Global_Var.index_Nd].value and GS_requirements == sheet[row][
                Global_Var.index_Service_Gov].value and dta_requirements == sheet[row][
                Global_Var.index_direction_to_action].value and fact_requirements == sheet[row][
                Global_Var.index_None].value:
                return row
        return None

    def create_pivot_table(self, dfs):
        pre_pivot_table = dfs.loc[
            (dfs["Напр.Деятельности"].isin(self.direction_do)) & (
                dfs["Группа направлений"].isin(self.group_direction))]
        values = ['Приход', 'Расход', pre_pivot_table.columns[6]]
        self.pivot_table = create_pivot_table(pre_pivot_table, 'КодСлужбыГС', values, 'sum')

    def automatic(self, dfs, templatePath):
        self.create_pivot_table(dfs)
        self.add_value_excel(templatePath)

    @cache
    def add_value_excel(self, templatePath):
        begin_row = Global_Var.start_revex
        for index in self.pivot_table.index:
            row = self.find_row(self.excel.sheet, "OPEX", index, "текущий запас", "факт", begin_row)
            begin_row = row
            if row is None:
                Global_Var.mistakes.append("REVEX " + str(index))
                begin_row = Global_Var.start_revex
                continue
            self.excel.push_cell(self.pivot_table, row, Global_Var.columns_reserve, index, self.pivot_table.columns[2])
            self.excel.push_cell(self.pivot_table, row, Global_Var.columns_profit, index, "Приход")
            self.excel.push_cell(self.pivot_table, row, Global_Var.columns_lost, index, "Расход")
        try:
            self.excel.workbook.save(templatePath)
        except:
            messagebox.showerror("Ошибка", "Нет доступа к файлу " + templatePath + " вероятно он открыт.")
        print("Successful enter Revex")
