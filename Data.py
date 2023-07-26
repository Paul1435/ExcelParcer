from datetime import datetime
import calendar
import locale
from dateutil.relativedelta import relativedelta


class Data():
    def __init__(self):
        print('data init')

    def common_format(self, data):
        locale.setlocale(locale.LC_ALL, 'ru_RU.UTF-8')
        date_obj = datetime.strptime(data, "%d.%m.%Y")
        return date_obj

    def table_format(self, date_obj):
        locale.setlocale(locale.LC_ALL, 'ru_RU.UTF-8')
        month = (str(calendar.month_name[date_obj.month])).lower()
        if month == "август" or month == "март":
            month = month + 'а'
        else:
            month = month[:-1] + 'я'
        print(f"{date_obj.day} {month} {date_obj.year}")
        return f"{date_obj.day} {month} {date_obj.year}"

    def current_time(self, data_pivot):
        cur_time = self.common_format(data_pivot)
        return self.table_format(cur_time)

    def next_time(self, data_obj):
        cur_time = self.common_format(data_obj)
        input_date = datetime.strptime(str(cur_time), '%Y-%m-%d %H:%M:%S')
        next_month_first_day = (input_date + relativedelta(months=1)).replace(day=1)
        return self.table_format(next_month_first_day)
