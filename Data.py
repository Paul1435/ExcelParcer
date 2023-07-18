from datetime import datetime, timedelta
import calendar
import locale
import pymorphy3


class Data():
    def __init__(self):
        self.changer = pymorphy3.MorphAnalyzer()

    def common_format(self, data):
        locale.setlocale(locale.LC_ALL, 'ru_RU.UTF-8')
        date_obj = datetime.strptime(data, "%d.%m.%Y")
        return date_obj

    def table_format(self, date_obj):
        locale.setlocale(locale.LC_ALL, 'ru_RU.UTF-8')
        parsed_month = pymorphy3.MorphAnalyzer().parse(calendar.month_name[date_obj.month])[0]
        month_name_decl = parsed_month.inflect({'sing', 'gent'}).word
        return f"{date_obj.day} {month_name_decl} {date_obj.year}"

    def current_time(self, data_pivot):
        cur_time = self.common_format(data_pivot)
        return self.table_format(cur_time)

    def next_time(self, data_obj):
        cur_time = self.common_format(data_obj)
        next_month = (cur_time + timedelta(days=int(data_obj[:2]))).replace(day=1)
        return self.table_format(next_month)
