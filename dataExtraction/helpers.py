def sanitizeValues(array):
    newArray = []
    for value in array:
        if value == "#DIV/0!":
            newArray = [*newArray, 0]
        else:
            newArray = [*newArray, value]
    return newArray


def sanitizePercentValues(array):
    newArray = []
    for value in array:
        if value == "#DIV/0!":
            newArray = [*newArray, 0]
        elif type(value) == str:
            newArray = [*newArray, value]
        else:
            newArray = [*newArray, round(value * 100, 2)]
    return newArray
