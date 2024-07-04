import xlwings as xw

excel_name = xw.Book("trial.xlsx")
sheet1 = excel_name.sheets["Sheet1"]

sheet1.range("A1:D1").options = "Welcome"
print(excel_name.sheets.active.name)
