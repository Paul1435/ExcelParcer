from functools import cache
from Pivot_Table import create_pivot_table
from tkinter import messagebox
import Global_Var


class Sz_Go_etc():
    def __init__(self, excel):
        self.pivot_table = None
        self.group_insurance_stock = ["Прочие, не учитываемые в расчете оборачиваемости"]
        self.excel = excel
        print("Инициализация ГоиЧс и СоиСИЗ")

    @cache
    def find_row(self, sheet, Nd_requirements, GS_requirements, dta_requirements, fact_requirements, begin_row):
        for row in range(begin_row, sheet.max_row + 1):
            if Nd_requirements == sheet[row][Global_Var.index_Nd].value and GS_requirements == sheet[row][
                Global_Var.index_Service_Gov].value and dta_requirements == sheet[row][
                Global_Var.index_direction_to_action].value and fact_requirements == sheet[row][
                Global_Var.index_None].value:
                return row
        return None

    def create_pivot_table(self, category):
        values = ['Приход', 'Расход', self.dictionary_pivot_table[category].columns[6]]
        self.pivot_table = create_pivot_table(self.dictionary_pivot_table[category], 'КодСлужбыГС', values, 'sum')

    def pre_pivot_table(self, dfs):
        self.dictionary_pivot_table = {"ЗАПАСЫ ГО": dfs.loc[
            (dfs["Группа направлений"].isin(self.group_insurance_stock)) & (
                dfs["Направление(Форма2)"].isin(["ГОиЧС"]))],
                                       "СОиСИЗ": dfs.loc[
                                           (dfs["Группа направлений"].isin(self.group_insurance_stock)) & (
                                               dfs["Направление(Форма2)"].isin(["СОиСИЗ"]))]}

    def automatic(self, dfs, templatePath, call_back):
        self.pre_pivot_table(dfs)
        for type in self.dictionary_pivot_table:
            self.create_pivot_table(type)
            self.add_value_excel(templatePath, type)
            Global_Var.step_load += 5
            call_back(Global_Var.step_load)
        print("Successful enter Го СОиСИЗ")

    def add_value_excel(self, templatePath, category):
        begin_row = Global_Var.start_etc
        for index in self.pivot_table.index:
            row = self.find_row(self.excel.sheet, category, index, "текущий запас", "факт", begin_row)
            if row is None:
                row = self.find_row(self.excel.sheet, None, index, "текущий запас", "факт", begin_row)
            begin_row = row
            if row is None:
                Global_Var.mistakes.append(category + " " + str(index))
                begin_row = Global_Var.start_etc
                continue
            self.excel.push_cell(self.pivot_table, row, Global_Var.columns_reserve, index,
                                 self.pivot_table.columns[2])
            self.excel.push_cell(self.pivot_table, row, Global_Var.columns_profit, index, "Приход")
            self.excel.push_cell(self.pivot_table, row, Global_Var.columns_lost, index, "Расход")
        try:
            self.excel.workbook.save(templatePath)
        except:
            messagebox.showerror("Ошибка", "Нет доступа к файлу " + templatePath + " вероятно он открыт.")
