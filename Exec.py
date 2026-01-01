import listSeen
import ramasSQL
import datetime


def main():
    listSeen.listSeen()
    startList = listSeen.startList
    multiArray = list()
    urlList = listSeen.urlList
    nameList = listSeen.nameList
    typeList = listSeen.typeList

    multiArray.append(nameList)
    multiArray.append(typeList)

    start = 0
    end = len(urlList)
    dateList = ""
    for i in range(start, end):
        if dateList == startList[i] or startList[i] == "\n":
            continue
        dateList = startList[i]
        arrDate = dateList.split("-")
        dateSee = datetime.date(year=2000 + int(arrDate[2]), day=int(arrDate[0]), month=int(arrDate[1]))
        urlname = urlList[i].replace("https://myanimelist.net", "")
        ramasSQL.ramas(multiArray, dateSee, urlList[i])
    ramasSQL.getEps()
    ramasSQL.excel("Excel\\Final_" + datetime.date.today().strftime("%d_%m_%Y") + ".xlsx")


main()
