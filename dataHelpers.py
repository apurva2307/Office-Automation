from datetime import datetime

currentMonth = datetime.now().month - 1
frMonth = currentMonth - 4 if currentMonth > 4 else currentMonth + 8

skip = [
    "PU3",
    "PU30",
    "PU42",
    "PU43",
    "PU44",
    "PU72",
    "PU73",
    "PU74",
    "PU75",
    "PU99",
    "NONSTAFF",
    "GROSS",
    "CREDIT",
    "NET",
]


def getMainData(monthdata, param):
    budget = monthdata[param]["budget"][-1]
    netTotal = monthdata[param]["toEndActuals"][-1]
    budgetUtil = monthdata[param]["budgetUtilization"][-1]
    varCPer = monthdata[param]["varAcCoppyPercent"][-1]
    return round(budget / 10000, 2), round(netTotal / 10000, 2), budgetUtil, varCPer


def highUtilStaff(monthdata, margin):
    result = {}
    for pu, value in monthdata.items():
        if pu == "STAFF":
            break
        if pu in skip:
            continue
        if (
            monthdata[pu]["budgetUtilization"][-1] > ((frMonth / 12) * 100) + margin
            and monthdata[pu]["toEndActuals"][-1] > 5000
        ):
            result[pu] = monthdata[pu]["budgetUtilization"][-1]
    return result


def highUtilNonStaff(monthdata, margin):
    result = {}
    for index, pu in enumerate(monthdata.keys()):
        staffIndex = list(monthdata.keys()).index("STAFF")
        if index > staffIndex:
            if pu in skip:
                continue
            if (
                monthdata[pu]["budgetUtilization"][-1] > ((frMonth / 12) * 100) + margin
                and monthdata[pu]["toEndActuals"][-1] > 5000
            ):
                result[pu] = monthdata[pu]["budgetUtilization"][-1]
    return result


def highUtilNonStaffCoppy(monthdata, margin):
    result = {}
    for index, pu in enumerate(monthdata.keys()):
        staffIndex = list(monthdata.keys()).index("STAFF")
        if index > staffIndex:
            if pu in skip:
                continue
            if (
                monthdata[pu]["varAcCoppyPercent"][-1] > margin
                and monthdata[pu]["toEndActuals"][-1] > 5000
            ):
                result[pu] = monthdata[pu]["varAcCoppyPercent"][-1]
    return result


def highUtilStaffCoppy(monthdata, margin):
    result = {}

    for pu, value in monthdata.items():
        if pu == "STAFF":
            break
        if pu in skip:
            continue
        if (
            monthdata[pu]["varAcCoppyPercent"][-1] > margin
            and monthdata[pu]["toEndActuals"][-1] > 5000
        ):
            result[pu] = monthdata[pu]["varAcCoppyPercent"][-1]
    return result


def highUtilNonStaffOther(dataOther, margin):
    result = {}
    keys = [
        "D-Traction",
        "E-Traction",
        "E-Office",
        "HSD-Civil",
        "HSD-Gen",
        "Lease",
        "IRFA",
        "Coach-C",
        "Station-C",
        "Colony-C",
    ]
    for key in keys:
        if dataOther[key][-1] > ((frMonth / 12) * 100) + margin:
            result[key] = dataOther[key][-1]
    return result
