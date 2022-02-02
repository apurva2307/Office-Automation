from data import extractData, extractDataSummary
from docx import Document
from datetime import datetime
from puToFullNameMap import puMap
from dataHelpers import *
from writingHelpers import *
from docx.enum.text import WD_ALIGN_PARAGRAPH

filePath = "OWE-DEC21.xlsx"
monthdata = extractData(filePath)
month = "Dec' 21"
budType = "RG"
marginExcessBud = 5
marginExcessCoppy = 20
currentMonth = datetime.now().month
frMonth = currentMonth - 4 if currentMonth > 4 else currentMonth + 8
puNameMap = puMap()
highUtilPuStaff = highUtilStaff(monthdata, marginExcessBud)
highUtilPuStaffCoppy = highUtilStaffCoppy(monthdata, marginExcessCoppy)
highUtilPuNonStaff = highUtilNonStaff(monthdata, marginExcessBud)
highUtilPuNonStaffCoppy = highUtilNonStaffCoppy(monthdata, marginExcessCoppy)
budget, netTotal, budgetUtil, varCPer = getMainData(monthdata, "NET")
staffBud, staffNet, staffUtil, staffVarC = getMainData(monthdata, "STAFF")
nonStaffBud, nonStaffNet, nonStaffUtil = (
    budget - staffBud,
    netTotal - staffNet,
    round(((netTotal - staffNet) * 100 / (budget - staffBud)), 2),
)
nonStaffCoppy = round(monthdata["NET"]["toEndActualsCoppy"][-1] / 10000, 2) - round(
    monthdata["STAFF"]["toEndActualsCoppy"][-1] / 10000, 2
)
nonStaffVarC = round(((nonStaffNet - nonStaffCoppy) * 100 / nonStaffCoppy), 2)
document = Document()
document.add_heading("NCR Financial Review DEC-2021", level=1)
document.add_heading("Revenue Expenditure:", level=1)
document.add_paragraph(
    f"The Revised Grant for Ord. Working Expenses (OWE) 2021-22, excluding suspense is Rs {budget} crore, more than SL by Rs. 430.10 crore and more than last year actuals by only Rs. 889.01 crore (11.37%).",
    style="List Bullet",
)
document.add_paragraph(
    f"OWE (excluding suspense) to end {month} amounts to Rs {netTotal} crore which is {budgetUtil}% of the {budType}, and {moreLess(varCPer)} COPPY by {varCPer}%.",
    style="List Bullet",
)
p1 = document.add_paragraph(
    f"Staff expenditure to end {month} of {staffNet} crore is {staffUtil}% of {budType}, and {moreLess(staffVarC)} COPPY by {staffVarC}%. Utilisation of {budType} is high for ",
    style="List Bullet",
)
iteratePara(p1, highUtilPuStaff)
p1.add_run(" Growth over last year is high for ")
iteratePara(p1, highUtilPuStaffCoppy)
p1.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
dataOther = extractDataSummary(filePath)
highUtilPuNonStaffOther = highUtilNonStaffOther(dataOther, marginExcessBud)
highUtilPuNonStaffOtherCoppy = highUtilNonStaffOther(dataOther, marginExcessCoppy)
p2 = document.add_paragraph(
    f"Non-Staff expenditure to end {month} of {nonStaffNet} crore is {nonStaffUtil}% of {budType} and {moreLess(nonStaffVarC)} COPPY by {nonStaffVarC}%. Utilisation of {budType} is high for ",
    style="List Bullet",
)
iterateParaSumm(p2, highUtilPuNonStaffOther)
iteratePara(p2, highUtilPuNonStaff)
p2.add_run(" Growth over last year is high for ")
iterateParaSumm(p2, highUtilPuNonStaffOtherCoppy)
iteratePara(p2, highUtilPuNonStaffCoppy)
p2.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
print(list(range(5, 8)))
document.save("FR_DEC21.docx")
