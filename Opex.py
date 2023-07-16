from datetime import timedelta
from functools import cache
from Data import Data
from Pivot_Table import create_pivot_table
from tkinter import messagebox
import Global_Var


class opex():
    def __init__(self, excel):
        self.pivot_table = None
        self.data = Data()
        self.class_est = [800, 802, 1800]
        self.group_direction = ["ОД_вспомогательные", "Основная деятельность"]
        self.filters_102_21 = set()
        self.filters_102_25 = set()
        self.direction_do = [5, 6, 60]
        self.excel = excel
        print("opex init")

    @cache
    def find_row(self, sheet, Nd_requirements, GS_requirements, dta_requirements, fact_requirements):
        for row in range(6, sheet.max_row + 1):
            if Nd_requirements == sheet[row][Global_Var.index_Nd].value and GS_requirements == sheet[row][
                Global_Var.index_Service_Gov].value and dta_requirements == sheet[row][
                Global_Var.index_direction_to_action].value and fact_requirements == sheet[row][
                Global_Var.index_None].value:
                return row
        return None

    def create_pivot_table(self, dfs):
        pre_pivot_table = dfs.loc[
            (~dfs["Напр.Деятельности"].isin(self.direction_do)) & (
                dfs["Группа направлений"].isin(self.group_direction)) & (
                ~dfs["Класс оценки"].isin(self.class_est))]
        values = ['Приход', 'Расход', pre_pivot_table.columns[6]]
        self.pivot_table = create_pivot_table(pre_pivot_table, 'КодСлужбыГС', values, 'sum')

    def automatic(self, dfs, templatePath):
        self.create_pivot_table(dfs)
        data_pivot = str(self.pivot_table.columns[2]).split(' ')[1]
        cur_time = self.current_time(data_pivot)
        need_period = self.next_time(data_pivot)
        self.add_value_excel(need_period, templatePath, cur_time)

    def current_time(self, data_pivot):
        cur_time = self.data.common_format(data_pivot)
        return self.data.table_format(cur_time)

    def next_time(self, data_pivot):
        cur_time = self.data.common_format(data_pivot)
        next_month = (cur_time + timedelta(days=int(data_pivot[:2]))).replace(day=1)
        return self.data.table_format(next_month)

    def add_value_excel(self, need_period, templatePath, cur_time):
        for index in self.pivot_table.index:
            row = self.find_row(self.excel.sheet, "OPEX", index, "текущий запас", "факт")
            columns_reserve = self.excel.find_column("Запасы на " + str(need_period), 5, 5)
            columns_profit = self.excel.find_column("Приход " + str(cur_time)[3:], 5, 5)
            columns_lost = []
            if len(columns_profit) != 0:
                columns_lost = self.excel.find_column("Расход " + str(cur_time)[3:], 5, 5)
            self.excel.push_cell(self.pivot_table, row, columns_reserve, index, self.pivot_table.columns[2])
            self.excel.push_cell(self.pivot_table, row, columns_profit, index, "Приход")
            self.excel.push_cell(self.pivot_table, row, columns_lost, index, "Расход")
        try:
            self.excel.workbook.save(templatePath)
        except:
            messagebox.showerror("Ошибка", "Нет доступа к файлу " + templatePath + " вероятно он открыт.")
        print("Successful enter opex")
