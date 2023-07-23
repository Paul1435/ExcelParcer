from tkinter import messagebox
from functools import cache
from Pivot_Table import create_pivot_table
import Global_Var


class cap_construction:
    def __init__(self, pushExcel):
        self.pivot_table = None
        self.direction_do = [6]
        self.excluded_prefix = '110203'
        self.exceptions = ['2770', '3034']
        self.excel = pushExcel
        self.winter_filter = set()
        self.other_filter = set()
        self.type = ["текущий запас", "ТЗБП", "ОП", "Ошибка"]
        print("init capcon")

    def delete_mistake(self):
        if "1020-11" in self.pivot_table.index:
            if '102-11' in self.pivot_table.index:
                self.pivot_table.loc['102-11'] += self.pivot_table.loc['1020-11']
            else:
                self.pivot_table.loc['102-11'] = 0
                self.pivot_table.loc['102-11'] += self.pivot_table.loc['1020-11']
            self.pivot_table = self.pivot_table.drop('1020-11')

    def general_table(self, dfs, type):
        if type != "ТЗБП":
            self.pivot_table = self.createPivotTable(dfs, type, self.other_filter)
            winter_pivot = self.createPivotTable(dfs, type, self.winter_filter)
            self.delete_mistake()
            if not self.pivot_table.empty:
                if not '102-11' in self.pivot_table.index:
                    self.pivot_table.loc['102-11'] = 0
                for index in winter_pivot.index:
                    self.pivot_table.loc['102-11'] += winter_pivot.loc[index]
            else:
                if not winter_pivot.empty:
                    self.pivot_table = winter_pivot
                    self.delete_mistake()

        if type == "ТЗБП":
            self.pivot_table = self.createPivotTable(dfs, type, self.other_filter)
            self.delete_mistake()

    def init_filters(self, dfs):
        for value in dfs['СПП-элемент']:
            str_val = str(value)
            if str_val.startswith(self.excluded_prefix) and (not str_val[-4:] in self.exceptions):
                self.winter_filter.add(value)
            if not str_val.startswith(self.excluded_prefix) or (
                    str_val.startswith(self.excluded_prefix) and str_val[-4:] in self.exceptions):
                self.other_filter.add(value)

    def createPivotTable(self, dfs, type, filtered_values):
        return create_pivot_table(self.pre_pivot_table(dfs, filtered_values)[type], 'КодСлужбыГС', self.values, 'sum')

    def pre_pivot_table(self, dfs, filtered_values):
        dictionary_pivot_table = {
            "текущий запас": dfs.loc[
                (dfs["Напр.Деятельности"].isin(self.direction_do)) & (dfs["СПП-элемент"].isin(filtered_values)) & (
                    dfs["Группа направлений"].isin(["ИД"]))],
            "ТЗБП": dfs.loc[
                (dfs["Направление(Форма2)"].isin(["ТЗБП"])) & (
                    dfs["Группа направлений"].isin(["Прочие, не учитываемые в расчете оборачиваемости"]))],
            "ОП": dfs.loc[
                (dfs["Напр.Деятельности"].isin(self.direction_do)) & (dfs["СПП-элемент"].isin(filtered_values)) & (
                    dfs["Группа направлений"].isin(["Опережающая поставка"]))],
            "Ошибка": dfs.loc[
                (dfs["Напр.Деятельности"].isin(self.direction_do)) & (dfs["СПП-элемент"].isin(filtered_values)) & (
                    dfs["Группа направлений"].isin(["Ошибка"]))]
        }
        return dictionary_pivot_table

    @cache
    def find_row(self, sheet, Nd_requirements, GS_requirements, dta_requirements, fact_requirements, row_begin):
        for row in range(row_begin, sheet.max_row + 1):
            if Nd_requirements == sheet[row][Global_Var.index_Nd].value and GS_requirements == sheet[row][
                Global_Var.index_Service_Gov].value and dta_requirements == sheet[row][
                Global_Var.index_direction_to_action].value and fact_requirements == sheet[row][
                Global_Var.index_None].value:
                return row
        return None

    @cache
    def add_value_excel(self, path, type):
        row_begin = Global_Var.start_cap_con

        if type != "ОП" and type != "Ошибка":
            for index in self.pivot_table.index:
                row = self.find_row(self.excel.sheet, "КС", index, type, "факт", row_begin)
                row_begin = row
                if row is None:
                    Global_Var.mistakes.append("КС " + str(type) + " " + str(index))
                    row_begin = Global_Var.start_cap_con
                    continue
                self.excel.push_cell(self.pivot_table, row, Global_Var.columns_reserve, index,
                                     self.pivot_table.columns[2])
                self.excel.push_cell(self.pivot_table, row, Global_Var.columns_profit, index, "Приход")
                self.excel.push_cell(self.pivot_table, row, Global_Var.columns_lost, index, "Расход")
        else:
            for index in self.pivot_table.index:
                row = self.find_row(self.excel.sheet, "КС", index, "текущий запас", "факт", row_begin)
                row_begin = row
                if row is None:
                    Global_Var.mistakes.append("КС " + str(type) + " " + str(index))
                    row_begin = Global_Var.start_cap_con
                    continue
                if type == "ОП":
                    self.excel.additional_res(self.pivot_table, row, [max(Global_Var.columns_reserve)], index,
                                              self.pivot_table.columns[2])
                    self.excel.push_cell(self.pivot_table, row, Global_Var.OP_column, index, "Приход")
                if type == "Ошибка":
                    self.excel.additional_res(self.pivot_table, row, Global_Var.columns_reserve, index,
                                              self.pivot_table.columns[2])
                    self.excel.additional_res(self.pivot_table, row, Global_Var.columns_profit, index, "Приход")
                    self.excel.additional_res(self.pivot_table, row, Global_Var.columns_lost, index, "Расход")
        try:
            self.excel.workbook.save(path)
        except:
            messagebox.showerror("Ошибка", "Нет доступа к файлу " + path + " вероятно он открыт.")
        print("Successful enter KC")

    def automatic(self, obj, template_obj, call_back):
        self.values = ['Приход', 'Расход', obj.columns[6]]
        self.init_filters(obj)
        for type in self.type:
            print(type)
            self.general_table(obj, type)
            print(self.pivot_table)
            self.add_value_excel(template_obj, type)
            Global_Var.step_load += 2
            call_back(Global_Var.step_load)
