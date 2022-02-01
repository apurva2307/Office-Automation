from puToFullNameMap import puMap

puNameMap = puMap()


def moreLess(data):
    return "more than" if data > 0 else "less than"


def iteratePara(p, data):
    for index, pu in enumerate(data.keys()):
        if index == len(data) - 1:
            p.add_run(f"and {puNameMap[pu]} ({data[pu]}%).")
        elif index == len(data) - 2:
            p.add_run(f"{puNameMap[pu]} ({data[pu]}%) ")
        else:
            p.add_run(f"{puNameMap[pu]} ({data[pu]}%), ")
