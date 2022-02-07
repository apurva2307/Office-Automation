from openpyxl import load_workbook
from dataExtraction.puList import getPUList, getPHs, getPHsMap
from dataExtraction.helpers import *
import requests, json
from decouple import config

TOKEN = config("TOKEN")


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
                detailedPuSheet.cell(i + 2, column).value,
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
    phs = getPHs()
    phsMap = getPHsMap()

    def phData(ph, rowRange, rowMap):
        for index, row in enumerate(rowRange):
            con = []
            open = []
            ncr = []
            for column in columns:
                if column < 6:
                    if column == 5:
                        con = [
                            *con,
                            sntzSigVPer(detailedPuSheet.cell(row, column).value),
                        ]
                        if ph == "EBR-P":
                            break
                    else:
                        con = [*con, sntzSigV(detailedPuSheet.cell(row, column).value)]
                if column < 10 and column > 6:
                    if column == 9:
                        open = [
                            *open,
                            sntzSigVPer(detailedPuSheet.cell(row, column).value),
                        ]
                    else:
                        open = [
                            *open,
                            sntzSigV(detailedPuSheet.cell(row, column).value),
                        ]
                if column < 14 and column > 10:
                    if column == 13:
                        ncr = [
                            *ncr,
                            sntzSigVPer(detailedPuSheet.cell(row, column).value),
                        ]
                    else:
                        ncr = [*ncr, sntzSigV(detailedPuSheet.cell(row, column).value)]
            if index == 0:
                if ph == "EBR-P":
                    result[f"{ph}"] = {f"{rowMap[index]}": {"NCR": con}}
                else:
                    result[f"{ph}"] = {
                        f"{rowMap[index]}": {"OPEN": open, "CON": con, "NCR": ncr}
                    }
            else:
                if ph == "EBR-P":
                    result[ph] = {
                        **result[ph],
                        f"{rowMap[index]}": {"NCR": con},
                    }
                else:
                    result[ph] = {
                        **result[ph],
                        f"{rowMap[index]}": {"OPEN": open, "CON": con, "NCR": ncr},
                    }

    for ph in phs:
        phData(ph, phsMap[ph]["rowRange"], phsMap[ph]["rowMap"])
    result["G-TOTAL"] = {
        "NCR": [
            sntzSigV(detailedPuSheet.cell(118, 11).value),
            sntzSigV(detailedPuSheet.cell(118, 12).value),
            sntzSigVPer(detailedPuSheet.cell(118, 13).value),
        ]
    }
    return result


def addToDatabase(month):
    registerURL = "https://e-commerce-api-apurva.herokuapp.com/api/v1/telebot/NCRAccountsBot/postData"
    data1 = extractData(f"./files/OWE-{month.upper()}.xlsx")
    payload = {
        "month": f"{month.upper()}",
        "type": "OWE",
        "data1": data1,
    }
    resp = requests.post(registerURL, json=payload)
    return resp.json()


def addToDatabaseCapex(filePath, sheet):
    registerURL = (
        "https://mydata.apurvasingh.dev/api/v1/telebot/NCRAccountsBot/postData"
    )
    data1 = extractDataCapex(filePath, sheet)
    payload = {
        "month": "JAN22",
        "type": "CAPEX",
        "data1": data1,
    }
    resp = requests.post(registerURL, json=payload)
    return resp.json()


def addToDatabaseCapexUpdate(filePath, sheet):
    registerURL = (
        "https://mydata.apurvasingh.dev/api/v1/telebot/NCRAccountsBot/updateData"
    )
    data1 = extractDataCapex(filePath, sheet)
    headers = {"token": TOKEN}
    payload = {
        "month": "JAN22",
        "type": "CAPEX",
        "data1": data1,
    }
    resp = requests.post(registerURL, json=payload, headers=headers)
    return resp.json()


def updateToDatabase(month):
    registerURL = (
        "https://mydata.apurvasingh.dev/api/v1/telebot/NCRAccountsBot/updateData"
    )
    data1 = extractData(f"../files/OWE-{month.upper()}.xlsx")
    headers = {"token": TOKEN}
    payload = {
        "month": f"{month.upper()}",
        "type": "OWE",
        "data1": data1,
    }
    resp = requests.post(registerURL, json=payload, headers=headers)
    return resp.json()


# if __name__ == "__main__":
#     data = addToDatabaseCapex("Capex Review 2021-22.xlsx", "Capex Jan-22")
#     print(data)
# print(data["EBR-IF"]["TOTAL"]["NCR"])
# if __name__ == "__main__":
#     # months = ["APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"]
#     # for month in months:
#     addToDatabase("JAN22")
#     print("done")
