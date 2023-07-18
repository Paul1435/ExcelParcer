from datetime import timedelta
from functools import cache
from Data import Data
from Pivot_Table import create_pivot_table
from tkinter import messagebox
import Global_Var


class drilling():
    def __init__(self, excel):
        self.pivot_table = None
        self.data = Data()
        self.filter = ["Агент", "Пропант", "Утяжелитель", "Песок"]
        self.filters_102_21 = set()
        self.filters_102_25 = set()
        self.direction_do = [5]
        self.excel = excel

    def create_pivot_table(self, dfs, filtered_values):
        pre_pivot_table = dfs.loc[
            (dfs["Напр.Деятельности"].isin(self.direction_do)) & (dfs["Кр. текст материала"].isin(filtered_values))]
        values = ['Приход', 'Расход', pre_pivot_table.columns[6]]
        return create_pivot_table(pre_pivot_table, 'КодСлужбыГС', values, 'sum')

    def general_table(self, dfs):
        self.pivot_table = self.create_pivot_table(dfs, self.filters_102_21)
        temp_table = self.create_pivot_table(dfs, self.filters_102_25)
        if len(self.pivot_table.index) > 1:
            self.pivot_table.loc['102-21'] += self.pivot_table.loc['102-25']
        self.pivot_table.loc['102-25'] = 0
        for index in temp_table.index:
            self.pivot_table.loc['102-25'] += temp_table.loc[index]

    #
    # def current_time(self, data_obj):
    #     cur_time = self.data.common_format(data_obj)
    #     return self.data.table_format(cur_time)
    #
    # def next_time(self, data_obj):
    #     cur_time = self.data.common_format(data_obj)
    #     next_month = (cur_time + timedelta(days=int(data_obj[:2]))).replace(day=1)
    #     return self.data.table_format(next_month)

    def create_filter(self, dfs):
        for value in dfs['Кр. текст материала']:
            str_val = str(value)
            if any(substr in str_val for substr in self.filter):
                self.filters_102_25.add(value)
            else:
                self.filters_102_21.add(value)

    @cache
    def find_row(self, sheet, Nd_requirements, GS_requirements, dta_requirements, fact_requirements, begin_row):
        for row in range(begin_row, sheet.max_row + 1):
            if Nd_requirements == sheet[row][Global_Var.index_Nd].value and GS_requirements == sheet[row][
                Global_Var.index_Service_Gov].value and dta_requirements == sheet[row][
                Global_Var.index_direction_to_action].value and fact_requirements == sheet[row][
                Global_Var.index_None].value:
                return row
        return None

    def add_value_excel(self, path):
        row_begin = Global_Var.start_drilling
        for index in self.pivot_table.index:
            row = self.find_row(self.excel.sheet, "Бурение", index, "текущий запас", "факт", row_begin)
            row_begin = row
            # columns_reserve = self.excel.find_column("Запасы на " + str(need_period), 5, 5)
            # columns_profit = self.excel.find_column("Приход " + str(cur_time)[3:], 5, 5)
            # columns_lost = []
            # if len(columns_profit) != 0:
            #     columns_lost = self.excel.find_column("Расход " + str(cur_time)[3:], 5, 5)
            self.excel.push_cell(self.pivot_table, row, Global_Var.columns_reserve, index, self.pivot_table.columns[2])
            self.excel.push_cell(self.pivot_table, row, Global_Var.columns_profit, index, "Приход")
            self.excel.push_cell(self.pivot_table, row, Global_Var.columns_lost, index, "Расход")
        try:
            self.excel.workbook.save(path)
        except:
            messagebox.showerror("Ошибка", "Нет доступа к файлу " + path + " вероятно он открыт.")
        print("Successful enter drilling")

    def automatic(self, obj, template_obj):
        self.create_filter(obj)
        self.general_table(obj)
        # data_pivot = str(self.pivot_table.columns[2]).split(' ')[1]
        # cur_time = self.current_time(data_pivot)
        # need_period = self.next_time(data_pivot)
        self.add_value_excel(template_obj)
