# To add a new cell, type '# %%'
# To add a new markdown cell, type '# %% [markdown]'
# %%
import pandas as pd
import numpy as np
import openpyxl

def auto_gen(menu_array, quantity_array):
    # %%
    table_source = np.zeros([10, 3])
    table_source[:, :] = np.nan
    dataframe = pd.DataFrame(table_source, columns=['商品名', '価格', '個数'])

    # %%

    for index, product_name in enumerate(menu_array):
        dataframe.loc[index, '商品名'] = product_name
    for index, quantity in enumerate(quantity_array):
        dataframe.loc[index, '個数'] = quantity
    dataframe

    # %%
    #エクセルファイル指定
    workbook = openpyxl.load_workbook('calculator.xlsx')
    #ワークシート指定
    sheet = workbook['販売個数']
    #最大行
    maxRow = sheet.max_row + 1

    col_name1 = 6
    col_name2 = 7
    for i in reversed(range(1, maxRow)):
        if sheet.cell(row=i, column=col_name1).value != None:
            for product_name, quantity in zip(dataframe['商品名'], dataframe['個数']):
                sheet.cell(i + 1, col_name1, value=product_name)
                sheet.cell(i + 1, col_name2, value=quantity)
                i = i + 1
            break

    workbook.save('calculator.xlsx')
