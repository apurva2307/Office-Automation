def sanitizeValues(array):
    newArray = []
    for value in array:
        if value == "#DIV/0!" or not value or value == "#REF!":
            newArray = [*newArray, 0]
        elif type(value) == str:
            newArray = [*newArray, value]
        else:
            newArray = [*newArray, round(value, 2)]
    return newArray


def sntzSigV(singleValue):
    if singleValue == "#DIV/0!" or singleValue == None or singleValue == "#REF!":
        return 0
    else:
        return round(singleValue, 2)


def sntzSigVPer(singleValue):
    if singleValue == "#DIV/0!" or singleValue == None or singleValue == "#REF!":
        return 0
    else:
        return round(singleValue * 100, 2)


def sanitizePercentValues(array):
    newArray = []
    for value in array:
        if value == "#DIV/0!" or value == "#REF!":
            newArray = [*newArray, 0]
        elif type(value) == str:
            newArray = [*newArray, value]
        else:
            newArray = [*newArray, round(value * 100, 2)]
    return newArray
