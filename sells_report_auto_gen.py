import pandas
import my_util
import get_date_info

def get_todays_menu(date_str)-> pandas.dataframe:
    file_path = '../output/sales_management.xlsx'
    datetime_date = my_util.get_datetime(date_str)
    just_date = datetime_date.date()

    dataframe = get_date_info.get_menu_info_bydate(date_str)

    return