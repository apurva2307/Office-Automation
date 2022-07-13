import pandas
import math

url = "https://docs.google.com/spreadsheets/d/1qNNWLcA75S1IekCQzvAHQhT_6aTyEph0kBAaEvVR0F4/gviz/tq?tqx=out:csv&sheet=Today'sFinalList"
url1 = "https://docs.google.com/spreadsheets/d/1qNNWLcA75S1IekCQzvAHQhT_6aTyEph0kBAaEvVR0F4/gviz/tq?tqx=out:csv&sheet=Nifty200"
url2 = "https://docs.google.com/spreadsheets/d/1qNNWLcA75S1IekCQzvAHQhT_6aTyEph0kBAaEvVR0F4/gviz/tq?tqx=out:csv&sheet=Midsmallcap400"
url3 = "https://docs.google.com/spreadsheets/d/1qNNWLcA75S1IekCQzvAHQhT_6aTyEph0kBAaEvVR0F4/gviz/tq?tqx=out:csv&sheet=PSUStocks"
data = pandas.read_csv(url)
listOfColumns = list(data)
newColNames = [
    "Nifty50",
    "Nifty50GTT",
    "ETF",
    "ETFGTT",
    "Nifty100",
    "Nifty100GTT",
    "PSU",
    "PSUGTT",
    "Nifty200",
    "Nifty200GTT",
    "NiftyMidSm",
    "NiftyMidSmGTT",
]
newCols = {}
for i, val in enumerate(listOfColumns):
    # printing a third element of column
    if i < 12:
        newCols[val] = newColNames[i]
    else:
        newCols[val] = "Ignore"

data.rename(columns=newCols, inplace=True)


def getData(data, type):
    result = {}
    for val1, val2 in zip(data[type].iteritems(), data[f"{type}GTT"].iteritems()):
        # if math.isnan (val2[1]):
        #     break
        if not isinstance(val1[1], str):
            break
        result[val1[1]] = val2[1]
    return result


nifty50 = getData(data, "Nifty50")
nifty100 = getData(data, "Nifty100")
psu = getData(data, "PSU")
niftymidsm = getData(data, "NiftyMidSm")
mainList = {}
mainList.update(nifty50)
mainList.update(nifty100)
mainList.update(psu)
mainList.update(niftymidsm)
print(mainList)
data2 = pandas.read_csv(url2)
listOfColumns2 = list(data2)
newCols2 = {}
for i, val in enumerate(listOfColumns2):
    if i == 0:
        newCols2[val] = "StockName"
    elif i == 2:
        newCols2[val] = "StockNameGTT"
data2.rename(columns=newCols2, inplace=True)
lookup = getData(data2, "StockName")
# print(lookup)


def changeInGtt(mainList, lookup):
    result = {}
    for stock, gtt in mainList.items():
        for stk, gttp in lookup.items():
            if stock == stk and gtt != gttp:
                result[stock] = gttp
    return result


print(changeInGtt(mainList, lookup))
