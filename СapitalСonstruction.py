from datetime import timedelta
from tkinter import messagebox
from functools import cache
from Data import Data
from Pivot_Table import create_pivot_table
import Global_Var


class cap_construction():
    def __init__(self, pushExcel):
        self.pivot_table = None
        self.data = Data()
        self.direction_do = [6]
        self.excluded_prefix = '110203'
        self.exceptions = ['2770', '3034']
        self.excel = pushExcel
        print("init capcon")

    def general_table(self, dfs):
        self.pivot_table = self.createPivotTable(dfs, self.cpp_no_winter_el_filter(dfs))
        winter_pivot = self.createPivotTable(dfs, self.cpp_winter_el_filter(dfs))
        if len(self.pivot_table.index) > 1:
            is_mistake_here = False
            for el in self.pivot_table.index:
                if el == "1020-11":
                    is_mistake_here = True
            if is_mistake_here:
                if '102-11' in self.pivot_table.index:
                    self.pivot_table.loc['102-11'] += self.pivot_table.loc['1020-11']
                else:
                    self.pivot_table.loc['102-11'] = 0
                    self.pivot_table.loc['102-11'] += self.pivot_table.loc['1020-11']
            self.pivot_table = self.pivot_table.drop('1020-11')

        else:
            self.pivot_table.loc['102-11'] = 0
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

    # def current_time(self, data_obj):
    #     cur_time = self.data.common_format(data_obj)
    #     return self.data.table_format(cur_time)
    #
    # def next_time(self, data_obj):
    #     cur_time = self.data.common_format(data_obj)
    #     next_month = (cur_time + timedelta(days=int(data_obj[:2]))).replace(day=1)
    #     return self.data.table_format(next_month)

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
    def add_value_excel(self, path):
        row_begin = Global_Var.start_cap_con
        for index in self.pivot_table.index:
            row = self.find_row(self.excel.sheet, "КС", index, "текущий запас", "факт", row_begin)
            row_begin = row
            # columns_reserve = self.excel.find_column("Запасы на " + str(need_period), 5, 5)
            # columns_profit = self.excel.find_column("Приход " + str(cur_time)[3:], 5, 5)
            # columns_lost = []
            # if len(columns_profit) != 0:
            #   columns_lost = self.excel.find_column("Расход " + str(cur_time)[3:], 5, 5)
            self.excel.push_cell(self.pivot_table, row, Global_Var.columns_reserve, index, self.pivot_table.columns[2])
            self.excel.push_cell(self.pivot_table, row, Global_Var.columns_profit, index, "Приход")
            self.excel.push_cell(self.pivot_table, row, Global_Var.columns_lost, index, "Расход")
        try:
            self.excel.workbook.save(path)
        except:
            messagebox.showerror("Ошибка", "Нет доступа к файлу " + path + " вероятно он открыт.")
        print("Successful enter KC")

    def automatic(self, obj, template_obj):
        self.general_table(obj)
        print(self.pivot_table)
        # data_pivot = str(self.pivot_table.columns[2]).split(' ')[1]
        #
        # cur_time = self.current_time(data_pivot)
        # need_period = self.next_time(data_pivot)
        # self.add_value_excel(need_period, template_obj, cur_time)
        self.add_value_excel(template_obj)
