from openpyxl import load_workbook
from dataExtraction.puList import getPUList
from dataExtraction.helpers import (
    sanitizeValues,
    sanitizePercentValues,
    sanitizeSingleValue,
)
import requests, json


def extractData(filePath):
    wb = load_workbook(filePath, data_only=True)
    detailedPuSheet = wb["PU Wise OWE"]
    # print(detailedPuSheet["B5"].value)
    # print(detailedPuSheet.cell(row=14, column=2).value)
    puList = getPUList()
    total = {}
    for column in range(3, len(puList) + 3):
        budget = []
        toEndBp = []
        toEndActualsCoppy = []
        toEndActuals = []
        varAcBp = []
        varAcBpPercent = []
        varAcCoppy = []
        varAcCoppyPercent = []
        budgetUtilization = []
        remainingBudget = []
        for i in range(5, 127, 11):
            budget = [*budget, detailedPuSheet.cell(row=i, column=column).value]
            toEndBp = [*toEndBp, detailedPuSheet.cell(row=i + 1, column=column).value]
            toEndActualsCoppy = [
                *toEndActualsCoppy,
                detailedPuSheet.cell(row=i + 2, column=column).value,
            ]
            toEndActuals = [
                *toEndActuals,
                detailedPuSheet.cell(row=i + 3, column=column).value,
            ]
            varAcBp = [*varAcBp, detailedPuSheet.cell(row=i + 4, column=column).value]
            varAcBpPercent = [
                *varAcBpPercent,
                detailedPuSheet.cell(row=i + 5, column=column).value,
            ]
            varAcCoppy = [
                *varAcCoppy,
                detailedPuSheet.cell(row=i + 6, column=column).value,
            ]
            varAcCoppyPercent = [
                *varAcCoppyPercent,
                detailedPuSheet.cell(row=i + 7, column=column).value,
            ]
            budgetUtilization = [
                *budgetUtilization,
                detailedPuSheet.cell(row=i + 8, column=column).value,
            ]
            remainingBudget = [
                *remainingBudget,
                detailedPuSheet.cell(row=i + 9, column=column).value,
            ]

        total[puList[column - 3]] = {
            "budget": sanitizeValues(budget),
            "toEndBp": sanitizeValues(toEndBp),
            "toEndActualsCoppy": sanitizeValues(toEndActualsCoppy),
            "toEndActuals": sanitizeValues(toEndActuals),
            "varAcBp": sanitizeValues(varAcBp),
            "varAcBpPercent": sanitizePercentValues(varAcBpPercent),
            "varAcCoppy": sanitizeValues(varAcCoppy),
            "varAcCoppyPercent": sanitizePercentValues(varAcCoppyPercent),
            "budgetUtilization": sanitizePercentValues(budgetUtilization),
            "remainingBudget": sanitizeValues(remainingBudget),
        }
    return total


def extractDataSummary(filePath):
    wb = load_workbook(filePath, data_only=True)
    detailedPuSheet = wb["Sheet1"]
    result = {}
    columns = [3, 6, 8, 9, 10, 11, 12, 13, 14, 15]
    columns1 = [3, 6, 8, 9, 11, 12, 13]
    rows = [5, 6, 42, 49, 60, 62, 63, 67, 109, 110, 111, 115, 116, 117]
    rowsMap = [
        "Staff",
        "Non-Staff",
        "D-Traction",
        "E-Traction",
        "E-Office",
        "HSD-Civil",
        "HSD-Gen",
        "Lease",
        "IRCA",
        "IRFA",
        "IRFC",
        "Coach-C",
        "Station-C",
        "Colony-C",
    ]
    for index, row in enumerate(rows):
        data = []
        if row < 88:
            for column in columns:
                if column == 12 or column == 14 or column == 15:
                    data = [
                        *data,
                        round((detailedPuSheet.cell(row, column).value) * 100, 2),
                    ]
                else:
                    data = [*data, round(detailedPuSheet.cell(row, column).value, 2)]
        else:
            for column in columns1:
                if column == 12 or column == 13:
                    data = [
                        *data,
                        round((detailedPuSheet.cell(row, column).value) * 100, 2),
                    ]
                else:
                    data = [*data, round(detailedPuSheet.cell(row, column).value, 2)]
        result[f"{rowsMap[index]}"] = data

    return result


def extractDataCapex(filePath, sheet):
    wb = load_workbook(filePath, data_only=True)
    detailedPuSheet = wb[sheet]
    result = {}
    columns = [3, 4, 5, 7, 8, 9, 11, 12, 13]
    phs = ["PH11", "PH14", "PH15"]
    phsMap = {
        "PH11": {
            "rowRange": list(range(5, 8)),
            "rowMap": ["CAP", "CAP(CH)", "SF", "TOTAL"],
        },
        "PH14": {
            "rowRange": list(range(10, 12)),
            "rowMap": ["CAP", "TOTAL"],
        },
        "PH15": {
            "rowRange": list(range(13, 16)),
            "rowMap": ["CAP", "CAP(RVNL)", "TOTAL"],
        },
    }

    def phData(ph, rowRange, rowMap):
        for index, row in enumerate(rowRange):
            data = []
            for column in columns:
                con = []
                if column < 6:
                    if column == 5:
                        con = [*con, detailedPuSheet.cell(row, column).value]
                    else:
                        con = [*con, detailedPuSheet.cell(row, column).value]
                open = []
                if column < 10:
                    if column == 9:
                        open = [*open, detailedPuSheet.cell(row, column).value]
                    else:
                        open = [*open, detailedPuSheet.cell(row, column)]
                ncr = []
                if column < 14:
                    if column == 13:
                        ncr = [*ncr, detailedPuSheet.cell(row, column).value]
                    else:
                        ncr = [*ncr, detailedPuSheet.cell(row, column).value]
            result[f"{ph}"] = {
                f"{rowMap[index]}": {"open": open, "con": con, "ncr": ncr}
            }

    for ph in phs:
        phData(ph, phsMap[ph]["rowRange"], phsMap[ph]["rowMap"])
    return result


def addToDatabase(month):
    registerURL = "https://e-commerce-api-apurva.herokuapp.com/api/v1/telebot/NCRAccountsBot/postData"
    data1 = extractData(f"../files/OWE-{month.upper()}.xlsx")
    payload = {
        "month": f"{month.upper()}",
        "type": "OWE",
        "data1": data1,
    }
    resp = requests.post(registerURL, json=payload)
    return resp.json()


def updateToDatabase(month):
    registerURL = "https://e-commerce-api-apurva.herokuapp.com/api/v1/telebot/NCRAccountsBot/updateData"
    data1 = extractData(f"../files/OWE-{month.upper()}.xlsx")
    headers = {"token": "Zr4u7x!A%C*F-JaNdRgUkXp2s5v8y/B?"}
    payload = {
        "month": f"{month.upper()}",
        "type": "OWE",
        "data1": data1,
    }
    resp = requests.post(registerURL, json=payload, headers=headers)
    return resp.json()


if __name__ == "__main__":
    print(extractDataCapex("Capex Review 2021-22.xlsx", "Capex Dec-21"))
    print("done")
# if __name__ == "__main__":
#     months = ["APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"]
#     for month in months:
#         updateToDatabase(f"{month}21")
#     print("done")
