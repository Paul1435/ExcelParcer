from datetime import timedelta
import openpyxl
from tkinter import messagebox
from functools import cache
from Data import Data
from Pivot_Table import create_pivot_table
from Push_Excel import push_excel


class cap_construction():
    def __init__(self):
        self.pivot_table = None
        self.data = Data()
        self.direction_do = [6]
        self.excluded_prefix = '110203'
        self.exceptions = ['2770', '3034']
        print("init capcon")

    def general_table(self, dfs):
        self.pivot_table = self.createPivotTable(dfs, self.cpp_no_winter_el_filter(dfs))
        winter_pivot = self.createPivotTable(dfs, self.cpp_winter_el_filter(dfs))
        is_mistake_here = False
        for el in self.pivot_table.index:
            if el == "1020-11":
                is_mistake_here = True
        if is_mistake_here:
            self.pivot_table.loc['102-11'] += self.pivot_table.loc['1020-11']
            self.pivot_table = self.pivot_table.drop('1020-11')

        for index in winter_pivot.index:
            self.pivot_table.loc['102-11'] += winter_pivot.loc[index]

    def cpp_winter_el_filter(self, dfs):
        filtered_values = set()
        for value in dfs['СПП-элемент']:
            str_val = str(value)
            if str_val.startswith(self.excluded_prefix) and (not str_val[-4:] in self.exceptions):
                filtered_values.add(value)
        return list(filtered_values)

    def cpp_no_winter_el_filter(self, dfs):
        filtered_values = set()
        for value in dfs['СПП-элемент']:
            str_val = str(value)
            if not str_val.startswith(self.excluded_prefix) or (
                    str_val.startswith(self.excluded_prefix) and str_val[-4:] in self.exceptions):
                filtered_values.add(value)
        return list(filtered_values)

    def createPivotTable(self, dfs, filtered_values):
        pre_pivot_table = dfs.loc[
            (dfs["Напр.Деятельности"].isin(self.direction_do)) & (dfs["СПП-элемент"].isin(filtered_values))]
        values = ['Приход', 'Расход', pre_pivot_table.columns[6]]
        return create_pivot_table(pre_pivot_table, 'КодСлужбыГС', values, 'sum')

    @cache
    def find_index(self, sheet, name):
        index = 0
        for cell in sheet.iter_cols(min_row=5, max_row=5, values_only=True):
            if cell[0] == name:
                return index
            index += 1

    def current_time(self, data_obj):
        cur_time = self.data.common_format(data_obj)
        return self.data.table_format(cur_time)

    def next_time(self, data_obj):
        cur_time = self.data.common_format(data_obj)
        next_month = (cur_time + timedelta(days=int(data_obj[:2]))).replace(day=1)
        return self.data.table_format(next_month)

    def find_const_index(self, sheet):
        self.index_Nd = self.find_index(sheet, "НД")
        self.index_Service_Gov = self.find_index(sheet, "Служба ГС")
        self.index_direction_to_action = self.find_index(sheet, "Направление деятельности")
        self.index_None = self.find_index(sheet, None)

    @cache
    def find_row(self, sheet, Nd_requirements, GS_requirements, dta_requirements, fact_requirements):
        for row in range(6, sheet.max_row + 1):
            if Nd_requirements == sheet[row][self.index_Nd].value and GS_requirements == sheet[row][
                self.index_Service_Gov].value and dta_requirements == sheet[row][
                self.index_direction_to_action].value and fact_requirements == sheet[row][self.index_None].value:
                return row
        return None

    @cache
    def add_value_excel(self, need_period, path, cur_time):
        excel = push_excel(path)
        self.find_const_index(excel.sheet)
        for index in self.pivot_table.index:
            row = self.find_row(excel.sheet, "КС", index, "текущий запас", "факт")
            columns_reserve = excel.find_column("Запасы на " + str(need_period), 5, 5)
            columns_profit = excel.find_column("Приход " + str(cur_time)[3:], 5, 5)
            columns_lost = []
            if len(columns_profit) != 0:
                columns_lost = excel.find_column("Расход " + str(cur_time)[3:], 5, 5)
            excel.push_cell(self.pivot_table, row, columns_reserve, index, self.pivot_table.columns[2])
            excel.push_cell(self.pivot_table, row, columns_profit, index, "Приход")
            excel.push_cell(self.pivot_table, row, columns_lost, index, "Расход")
        try:
            excel.workbook.save(path)
        except:
            messagebox.showerror("Ошибка", "Нет доступа к файлу " + path + " вероятно он открыт.")
        print("Successful enter KC")

    def automatic(self, obj, template_obj):
        self.general_table(obj)
        data_pivot = str(self.pivot_table.columns[2]).split(' ')[1]
        cur_time = self.current_time(data_pivot)
        need_period = self.next_time(data_pivot)
        self.add_value_excel(need_period, template_obj, cur_time)
