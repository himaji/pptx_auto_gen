import pandas

data_path = "calculator.xlsx"

data = pandas.read_excel(data_path, sheet_name =None)



print(data)