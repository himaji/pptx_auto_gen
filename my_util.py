# To add a new cell, type '# %%'
# To add a new markdown cell, type '# %% [markdown]'
# %%
import datetime
import locale

def get_datetime(date_str) -> datetime:
    date_splited = date_str.split('/')

    # %%
    formated_date = datetime.datetime.strptime(
        '2021/' + date_splited[0] + '/' + date_splited[1], "%Y/%m/%d")

    return formated_date