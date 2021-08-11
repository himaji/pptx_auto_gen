import pandas
import my_util
import linking_pandas_and_excel

def get_todays_menu(date_str)-> pandas.core.frame.DataFrame:
    file_path = '../output/sales_management.xlsx'
    datetime_date = my_util.get_datetime(date_str)
    just_date = datetime_date.date()

    dataframe = linking_pandas_and_excel.get_data_bydate(date_str)

    return