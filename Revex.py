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

    def pre_pivot_table(self, dfs):
        self.dictionary_pivot_table = {
            "текущий запас": dfs.loc[
                (dfs["Напр.Деятельности"].isin(self.direction_do)) & (
                    dfs["Группа направлений"].isin(self.group_direction))],
            "страховые запасы": dfs.loc[
                (dfs["Напр.Деятельности"].isin(self.direction_do)) & (
                    dfs["Группа направлений"].isin(["Прочие, не учитываемые в расчете оборачиваемости"])) & dfs[
                    "Категория запаса"].isin(["SZ"])],
            "НВИ": dfs.loc[
                (dfs["Напр.Деятельности"].isin(self.direction_do)) & (
                    dfs["Группа направлений"].isin(["Прочие, учитываемые в расчете оборачиваемости"])) & dfs[
                    "Категория запаса"].isin(["NV"])],
            "ОП": dfs.loc[
                (dfs["Напр.Деятельности"].isin(self.direction_do)) & (
                    dfs["Группа направлений"].isin(["Опережающая поставка"]))],
            "Ошибка": dfs.loc[
                (dfs["Напр.Деятельности"].isin(self.direction_do)) & (
                    dfs["Группа направлений"].isin(["Ошибка"]))]
        }

    def create_pivot_table(self, dfs, category):
        values = ['Приход', 'Расход', self.dictionary_pivot_table[category].columns[6]]
        self.pivot_table = create_pivot_table(self.dictionary_pivot_table[category], 'КодСлужбыГС', values, 'sum')

    def automatic(self, dfs, templatePath, call_back):
        self.pre_pivot_table(dfs)
        for category in self.dictionary_pivot_table:
            print(category)
            self.create_pivot_table(dfs, category)
            print(self.pivot_table)
            self.add_value_excel(templatePath, category)
            Global_Var.step_load += 2
            call_back(Global_Var.step_load)

    @cache
    def add_value_excel(self, templatePath, category):
        begin_row = Global_Var.start_revex
        for index in self.pivot_table.index:
            if category == "страховые запасы" or category == "НВИ":
                row = self.find_row(self.excel.sheet, "OPEX", index, category, "факт", begin_row)
            else:
                row = self.find_row(self.excel.sheet, "OPEX", index, "текущий запас", "факт", begin_row)
            if row is None:
                if category == "страховые запасы" or category == "НВИ":
                    row = self.find_row(self.excel.sheet, "OPEX", index, category, "факт", Global_Var.start_revex)
                else:
                    row = self.find_row(self.excel.sheet, "OPEX", index, "текущий запас", "факт",
                                        Global_Var.start_revex)
            begin_row = row
            print(row)
            if row is None:
                Global_Var.mistakes.append("REVEX " + str(index))
                begin_row = Global_Var.start_revex
                continue
            if category == "Ошибка":
                self.excel.additional_res(self.pivot_table, row, Global_Var.columns_reserve, index,
                                          self.pivot_table.columns[2])
                self.excel.additional_res(self.pivot_table, row, Global_Var.columns_profit, index, "Приход")
                self.excel.additional_res(self.pivot_table, row, Global_Var.columns_lost, index, "Расход")
            elif category == "ОП":
                self.excel.additional_res(self.pivot_table, row, [max(Global_Var.columns_reserve)], index,
                                          self.pivot_table.columns[2])
                self.excel.additional_res(self.pivot_table, row, Global_Var.OP_column, index, "Приход")
            else:
                self.excel.additional_res(self.pivot_table, row, Global_Var.columns_reserve, index,
                                          self.pivot_table.columns[2])
                self.excel.additional_res(self.pivot_table, row, Global_Var.columns_profit, index, "Приход")
                self.excel.additional_res(self.pivot_table, row, Global_Var.columns_lost, index, "Расход")
        try:
            self.excel.workbook.save(templatePath)
        except:
            messagebox.showerror("Ошибка", "Нет доступа к файлу " + templatePath + " вероятно он открыт.")
        print("Successful enter Revex")
