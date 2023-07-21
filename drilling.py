from functools import cache
from Pivot_Table import create_pivot_table
from tkinter import messagebox
import Global_Var


class drilling():
    def __init__(self, excel):
        self.pivot_table = None
        self.filter = ["Агент", "Пропант", "Утяжелитель", "Песок"]
        self.filters_102_21 = set()
        self.filters_102_25 = set()
        self.direction_do = [5]
        self.excel = excel

    def create_pivot_table(self, dfs, filtered_values, type):
        dictionary_pivot_table = {
            "текущий запас": dfs.loc[
                (dfs["Напр.Деятельности"].isin(self.direction_do)) & (
                    dfs["Кр. текст материала"].isin(filtered_values))],
            "ОП": dfs.loc[(dfs["Группа направлений"].isin(["Опережающая поставка"])) &
                          (dfs["Напр.Деятельности"].isin(self.direction_do)) & (
                              dfs["Кр. текст материала"].isin(filtered_values))]
        }
        values = ['Приход', 'Расход', dfs.columns[6]]
        return create_pivot_table(dictionary_pivot_table[type], 'КодСлужбыГС', values, 'sum')

    def general_table(self, dfs, type):
        self.pivot_table = self.create_pivot_table(dfs, self.filters_102_21, type)
        temp_table = self.create_pivot_table(dfs, self.filters_102_25, type)
        if len(self.pivot_table.index) > 1:
            self.pivot_table.loc['102-21'] += self.pivot_table.loc['102-25']
        self.pivot_table.loc['102-25'] = 0
        if len(temp_table.index) == 0:
            self.pivot_table.drop('102-25')
        for index in temp_table.index:
            self.pivot_table.loc['102-25'] += temp_table.loc[index]

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

    def add_value_excel(self, path, category):
        row_begin = Global_Var.start_drilling
        for index in self.pivot_table.index:
            row = self.find_row(self.excel.sheet, "Бурение", index, "текущий запас", "факт", row_begin)
            row_begin = row
            if row is None:
                Global_Var.mistakes.append("Бурение " + str(index))
                row_begin = Global_Var.start_drilling
                continue
            if (category == "ОП"):
                self.excel.additional_res(self.pivot_table, row, [max(Global_Var.columns_reserve)], index,
                                          self.pivot_table.columns[2])
                self.excel.push_cell(self.pivot_table, row, Global_Var.OP_column, index, "Приход")
            else:
                self.excel.push_cell(self.pivot_table, row, Global_Var.columns_reserve, index,
                                     self.pivot_table.columns[2])
                self.excel.push_cell(self.pivot_table, row, Global_Var.columns_profit, index, "Приход")
                self.excel.push_cell(self.pivot_table, row, Global_Var.columns_lost, index, "Расход")
        try:
            self.excel.workbook.save(path)
        except:
            messagebox.showerror("Ошибка", "Нет доступа к файлу " + path + " вероятно он открыт.")
        print("Successful enter drilling")

    def automatic(self, obj, template_obj, call_back):
        type = ["текущий запас", 'ОП']
        self.create_filter(obj)
        for category in type:
            print(category)
            self.general_table(obj, category)
            print(self.pivot_table)
            self.add_value_excel(template_obj, category)
            Global_Var.step_load += 4
            call_back(Global_Var.step_load)
