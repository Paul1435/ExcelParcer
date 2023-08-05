from functools import cache
from Pivot_Table import create_pivot_table
from tkinter import messagebox
import Global_Var


class OHCC():
    def __init__(self, excel):
        self.pivot_table = None
        self.class_est = [800, 802, 1800]
        self.group_direction = ["ОНСС", "ОНСС_вспомогательные"]
        self.direction_do = [5, 6, 60]
        self.excel = excel
        print("OHCC init")

    def pre_pivot_table(self, dfs):
        self.dictionary_pivot_table = {
            "текущий запас": dfs.loc[
                (~dfs["Напр.Деятельности"].isin(self.direction_do)) & (
                    dfs["Группа направлений"].isin(self.group_direction)) & (
                    dfs["Класс оценки"].isin(self.class_est))],
            "страховые запасы": dfs.loc[(~dfs["Напр.Деятельности"].isin(self.direction_do)) &
                                        (dfs["Направление(Форма2)"].isin(["Страховые запасы и Аварийные запасы"])) & (
                                            dfs["Группа направлений"].isin(
                                                ["Прочие, не учитываемые в расчете оборачиваемости"])) & (
                                            dfs["Класс оценки"].isin(self.class_est))],

            "вторичное сырье": dfs.loc[(~dfs["Напр.Деятельности"].isin(self.direction_do)) &
                                       (dfs["Направление(Форма2)"].isin(["Втор. сырье"])) & (
                                           dfs["Группа направлений"].isin(
                                               ["Прочие, учитываемые в расчете оборачиваемости"])) & (
                                           dfs["Класс оценки"].isin(self.class_est))],

            "НВИ": dfs.loc[(~dfs["Напр.Деятельности"].isin(self.direction_do)) &
                           (dfs["Направление(Форма2)"].isin(["НВИ/НЛИ"])) & (
                               dfs["Группа направлений"].isin(["Прочие, учитываемые в расчете оборачиваемости"])) & (
                               dfs["Класс оценки"].isin(self.class_est)) & (
                               dfs["Категория запаса"].isin(["NV"]))],

            "НЛИ": dfs.loc[(~dfs["Напр.Деятельности"].isin(self.direction_do)) &
                           (dfs["Направление(Форма2)"].isin(["НВИ/НЛИ"])) & (
                               dfs["Группа направлений"].isin(["Прочие, учитываемые в расчете оборачиваемости"])) & (
                               dfs["Класс оценки"].isin(self.class_est)) & (
                               dfs["Категория запаса"].isin(["NL"]))],
            "ОП": dfs.loc[
                (~dfs["Напр.Деятельности"].isin(self.direction_do)) & (
                    dfs["Группа направлений"].isin(["Опережающая поставка"])) & (
                    dfs["Класс оценки"].isin(self.class_est))],
            "Ошибка": dfs.loc[
                (~dfs["Напр.Деятельности"].isin(self.direction_do)) & (
                    dfs["Группа направлений"].isin(["Ошибка"])) & (
                    dfs["Класс оценки"].isin(self.class_est))]
        }

    def automatic(self, dfs, templatePath, call_back):
        self.pre_pivot_table(dfs)
        for category in self.dictionary_pivot_table:
            print(category)
            self.create_pivot_table(category)
            print(self.pivot_table)
            self.add_value_excel(templatePath, category)
            Global_Var.step_load += 3
            call_back(Global_Var.step_load)
        print("Заполнение ОНСС закончено")

    def create_pivot_table(self, category):
        values = ['Приход', 'Расход', self.dictionary_pivot_table[category].columns[6]]
        self.pivot_table = create_pivot_table(self.dictionary_pivot_table[category], 'КодСлужбыГС', values, 'sum')
        if "1020-11" in self.pivot_table.index:
            if '102-11' in self.pivot_table.index:
                self.pivot_table.loc['102-11'] += self.pivot_table.loc['1020-11']
            else:
                self.pivot_table.loc['102-11'] = 0
                self.pivot_table.loc['102-11'] += self.pivot_table.loc['1020-11']
            self.pivot_table = self.pivot_table.drop('1020-11')

    def add_value_excel(self, templatePath, category):
        begin_row = Global_Var.start_equipment
        sub_category = category
        if category == "ОП" or category == "Ошибка":
            sub_category = "текущий запас"
        for index in self.pivot_table.index:
            if (index == "102-04" or index == "102-11"):
                row = self.find_row(self.excel.sheet, "КС", index, sub_category, "факт", Global_Var.start_cap_con)
                begin_row = Global_Var.start_equipment
            else:
                row = self.find_row(self.excel.sheet, "ОНСС", index, sub_category, "факт", begin_row)
                begin_row = row
            if row is None:
                begin_row = Global_Var.start_equipment
                row = self.find_row(self.excel.sheet, "ОНСС", index, sub_category, "факт", begin_row)
                begin_row = row
            if row is None:
                Global_Var.mistakes.append("ОНСС " + str(category) + " " + str(index))
                begin_row = Global_Var.start_equipment
                continue
            if category == "Ошибка":
                self.excel.additional_res(self.pivot_table, row, Global_Var.columns_reserve, index,
                                          self.pivot_table.columns[2])
                self.excel.additional_res(self.pivot_table, row, Global_Var.columns_profit, index, "Приход")
                self.excel.additional_res(self.pivot_table, row, Global_Var.columns_lost, index, "Расход")
            elif category != "ОП":
                self.excel.additional_res(self.pivot_table, row, Global_Var.columns_reserve, index,
                                          self.pivot_table.columns[2])
                self.excel.additional_res(self.pivot_table, row, Global_Var.columns_profit, index, "Приход")
                self.excel.additional_res(self.pivot_table, row, Global_Var.columns_lost, index, "Расход")
            else:
                self.excel.additional_res(self.pivot_table, row, [max(Global_Var.columns_reserve)], index,
                                          self.pivot_table.columns[2])
                self.excel.additional_res(self.pivot_table, row, Global_Var.OP_column, index, "Приход")
        try:
            self.excel.workbook.save(templatePath)
        except:
            messagebox.showerror("Ошибка", "Нет доступа к файлу " + templatePath + " вероятно он открыт.")

    @cache
    def find_row(self, sheet, Nd_requirements, GS_requirements, dta_requirements, fact_requirements, row_begin):
        for row in range(row_begin, sheet.max_row + 1):
            if Nd_requirements == sheet[row][Global_Var.index_Nd].value and GS_requirements == sheet[row][
                Global_Var.index_Service_Gov].value and dta_requirements == sheet[row][
                Global_Var.index_direction_to_action].value and fact_requirements == sheet[row][
                Global_Var.index_None].value:
                return row
        return None
