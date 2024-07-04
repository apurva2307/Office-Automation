import xlwings as xw
import pandas as pd

filepath = "files\BudgetReport.xls"
# excel_name = xw.Book(filepath)
# sheet1 = excel_name.sheets[" DPW_DW_HW_SW_PW "]
# print(sheet1.range("A5:V11784").value)
df = pd.read_excel(filepath, header=4, sheet_name=" DPW_DW_HW_SW_PW ")
print(df.head())
