from datetime import datetime

currentMonth = datetime.now().month
frMonth = currentMonth - 4 if currentMonth > 4 else currentMonth + 8

skip = [
    "PU3",
    "PU7",
    "PU9",
    "PU14",
    "PU15",
    "PU17",
    "PU19",
    "PU20",
    "PU22",
    "PU23",
    "PU24",
    "PU25",
    "PU29",
    "PU30",
    "PU31",
    "PU33",
    "PU36",
    "PU37",
    "PU38",
    "PU40",
    "PU41",
    "PU42",
    "PU43",
    "PU48",
    "PU51",
    "PU52",
    "PU53",
    "PU44",
    "PU60",
    "PU61",
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


def inCrore(value):
    return round(value / 10000, 2)


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
        "D-TRACTION",
        "E-TRACTION",
        "E-OFFICE",
        "HSD-CIVIL",
        "HSD-GEN",
        "LEASE",
        "IRFA",
        "COACH-C",
        "STATION-C",
        "COLONY-C",
    ]
    for key in keys:
        if dataOther[key][-1] > ((frMonth / 12) * 100) + margin:
            result[key] = dataOther[key][-1]
    return result


def highUtilNonStaffOtherCoppy(dataOther, margin):
    result = {}
    keys = [
        "D-TRACTION",
        "E-TRACTION",
        "E-OFFICE",
        "HSD-CIVIL",
        "HSD-GEN",
        "LEASE",
        "IRFA",
        "COACH-C",
        "STATION-C",
        "COLONY-C",
    ]
    for key in keys:
        if dataOther[key][-2] > ((frMonth / 12) * 100) + margin:
            result[key] = dataOther[key][-2]
    return result


def slowProgCapex(dataCapex, sof, margin):
    result = {}
    for key in dataCapex.keys():
        if key in ["TOTAL", "G-TOTAL", "EBR-IF", "EBR-P"]:
            continue
        else:
            if sof in dataCapex[key].keys():
                if dataCapex[key][sof]["NCR"][-1] < ((frMonth / 12) * 100) - margin:
                    result[key] = dataCapex[key][sof]["NCR"][-1]
    return result


def highProgCapex(dataCapex, sof, margin):
    result = {}
    for key in dataCapex.keys():
        if key in ["TOTAL", "G-TOTAL", "EBR-IF", "EBR-P"]:
            continue
        else:
            if sof in dataCapex[key].keys():
                if dataCapex[key][sof]["NCR"][-1] > ((frMonth / 12) * 100) + margin:
                    result[key] = dataCapex[key][sof]["NCR"][-1]
    return result
