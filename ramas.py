import re
import os
import datetime
import calendar
import time
import requests
import openpyxl
from bs4 import BeautifulSoup
import pandas as pd
from openpyxl.styles import NamedStyle, Alignment



def getList(name):
    my_file = open(baseTxt + name, "r", encoding="utf-8")
    data = my_file.read()
    my_file.close()
    return data.split("\n")


def urllink(url):
    try:
        return requests.get(url)
    except:
        return urllink(url)


def convertDate(timestr, dateStart):
    if timestr == "Not available":
        return datetime.date(year=2099, day=31, month=12)
    else:
        if timestr == "?":
            return dateStart

    arrDate = timestr.split(" ")
    tam = len(arrDate)
    switcher = {"Jan": 1, "Feb": 2, "Mar": 3, "Apr": 4, "May": 5, "Jun": 6,
                "Jul": 7, "Aug": 8, "Sep": 9, "Oct": 10, "Nov": 11, "Dec": 12}

    year = 0
    month = 0
    day = 0

    if tam == 1:
        year = int(arrDate[0])
        month = 12
        day = 31
    if tam == 2:
        year = int(arrDate[1])
        month = switcher.get(arrDate[0], 0)
        day = calendar.monthrange(year, month)[1]
    if tam == 3:
        year = int(arrDate[2])
        month = switcher.get(arrDate[0], 0)
        day = int(arrDate[1])

    date = datetime.date(year=year, day=day, month=month)
    return date


def ramas(listaStatus, date_x, link):
    global err
    page = urllink(link)
    soup = BeautifulSoup(page.content.decode('utf-8'), 'html.parser')
    try:
        name = soup.find_all('h1', class_='title-name h1_bold_none')[0].text
    except:
        print("Error:" + str(err) + " " + link)
        if link == "/anime//-":
            print(listTitle[len(listTitle) - 1])
            return
        if link == "/anime//":
            return
        err = err + 1
        time.sleep(5 * 60)
        ramas(listaStatus, date_x, link)
        return

    urlList.append(link)
    listTitle.append(name)
    seenList.append(date_x)

    opnening = soup.find_all('div', class_='opnening')
    allTr = opnening[0].find_all('table')[1].find_all('tr')
    opname = ""
    for tr in allTr:
        allTd = tr.find_all('td')
        if len(allTd) > 1:
            if opname != "":
                opname = opname + "\n"
            opname = opname + allTd[1].extract().text.rstrip().lstrip()
    opList.append(opname)

    ending = soup.find_all('div', class_='ending')
    allTr = ending[0].find_all('table')[0].find_all('tr')
    endname = ""
    for tr in allTr:
        allTd = tr.find_all('td')
        if len(allTd) > 1:
            if endname != "":
                endname = endname + "\n"
            endname = endname + allTd[1].extract().text.rstrip().lstrip()
    endList.append(endname)

    statusName = ""
    name2 = name
    for j in letterProhibited:
        name2 = name2.replace(j, "")

    listTitleWithouthEspecialChars.append(name2)
    print(name2)

    if len(listaStatus) > 0:
        try:
            index = listaStatus[0].index(name2)
            statusName = listaStatus[1][index]
        except:
            pass
    statusList.append(statusName)

    listTd = soup.find_all('td', class_='borderClass')
    h2List = listTd[0].find_all('h2')
    i = 0
    h2 = []
    for j in h2List:
        if j.text == "Information":
            h2 = h2List[i]
        i = i + 1

    spans = h2.find_next_siblings()
    for k in spans:
        for t in k.select('span'):
            t.extract()
            if t.text == "Type:":
                typeList.append(k.text.rstrip().lstrip())
            if t.text == "Studios:":
                studiosList.append(k.text.rstrip().lstrip())
            if t.text == "Duration:":
                durationList.append(k.text.rstrip().lstrip())
            if t.text == "Episodes:":
                episodesString = k.text.rstrip().lstrip()
                if episodesString != "Unknown":
                    episodeList.append(int(episodesString))
                else:
                    episodeList.append(999)
            if t.text == "Aired:":
                textdate = k.text.replace(",", "")
                arrDate = textdate.split("to")
                textdate = arrDate[0].rstrip().lstrip()
                numberdate = convertDate(textdate, datetime.date(year=2099, day=31, month=12))
                if len(arrDate) > 1:
                    numberdateFinish = convertDate(arrDate[1].rstrip().lstrip(), numberdate)
                else:
                    numberdateFinish = numberdate
                dateList.append(numberdate)
                numberdateFinishList.append(numberdateFinish)

    if link in fromUrl:
        link = toUrl[fromUrl.index(link)]
        ramas(listaStatus, date_x, link)

    if name in nameProhibitedIncluded:
        return

    tableGeneral = soup.find_all("table", class_=re.compile("entries-table"))
    for table in tableGeneral:
        tdList = table.find_all('td', class_='fw-n')
        for td in tdList:
            text = td.text.rstrip().lstrip().replace('\n                  ', ' ')
            if text == "Character:":
                print(name)

            # if not "Piece" in name:
            #    print(name)

            if (text in listProhibited) or \
                    (text in listProhibited2 and name in characterNotPermited) or \
                    (text in categoryNotPermited and name in nameCategoryNotPermited) or \
                    (text in categoryNotPermited2 and name in nameCategoryNotPermited2) or \
                    (text in categoryNotPermited3 and name in nameCategoryNotPermited3) or \
                    (text in categoryNotPermited4 and name in nameCategoryNotPermited4):
                continue
            else:
                alink = td.find_next_siblings("td")[0].find_all("a", href=True)
                for a in alink:
                    if not (a.text.rstrip().lstrip().replace("‚òÖ", "★") in listTitle) and not (a.text in nameProhibited):
                        if not (text in categoryNotPermited3 and a.text in nameSpecialProhibited):
                            ramas(listaStatus, date_x, a['href'])

    tableGeneral = soup.find_all("div", class_=re.compile("entries-tile"))
    for table in tableGeneral:
        tdList = table.find_all('div', class_='relation')
        for td in tdList:
            text = td.text.rstrip().lstrip().replace('\n                  ', ' ')
            if text == "Character:":
                print(name)

            # if not "Piece" in name:
            #    print(name)

            if (text in listProhibited) or \
                    (text in listProhibited2 and name in characterNotPermited) or \
                    (text in categoryNotPermited and name in nameCategoryNotPermited) or \
                    (text in categoryNotPermited2 and name in nameCategoryNotPermited2) or \
                    (text in categoryNotPermited3 and name in nameCategoryNotPermited3) or \
                    (text in categoryNotPermited4 and name in nameCategoryNotPermited4):
                continue
            else:
                alink = td.find_next_siblings("div")[0].find_all("a", href=True)
                for a in alink:
                    if not (a.text.rstrip().lstrip().replace("‚òÖ", "★") in listTitle) and not (a.text in nameProhibited):
                        if not (text in categoryNotPermited3 and a.text in nameSpecialProhibited):
                            ramas(listaStatus, date_x, a['href'])


def getEps():
    global df
    capList = [None] * len(urlList)
    df = pd.DataFrame({
        'Nombre': listTitleWithouthEspecialChars,
        'Fecha Visto': seenList,
        'Fecha Inicio': dateList,
        'Fecha Fin': numberdateFinishList,
        'Tipo': typeList,
        'Episodios': episodeList,
        'Duración': durationList,
        'Cap. Sueltos': capList,
        'Studios': studiosList,
        'Opening': opList,
        'Ending': endList,
        'Estado': statusList,
        'Url': urlList,

    })
    df.sort_values(by=['Fecha Visto', 'Fecha Inicio', 'Nombre'], ascending=[False, True, False], inplace=True)

    seenDate = datetime.date(1900, 1, 1)
    daysBetween = 0
    i = 0
    prevMax = 0
    dateStart = datetime.date(1900, 1, 1)
    dateEnd = datetime.date(1900, 1, 1)
    for ele in df['Fecha Visto']:
        if ele != seenDate:
            dateStart = datetime.date(1900, 1, 1)
            dateEnd = datetime.date(1900, 1, 1)
            seenDate = ele
            prevMax = i

        dateCellS = df.iloc[i]['Fecha Inicio']
        dateCellE = df.iloc[i]['Fecha Fin']
        pass_x = False
        if dateCellS < dateEnd:
            calculateDays(dateCellS, dateStart, daysBetween, dateCellE, i, prevMax)
        else:
            for k in range(prevMax, i):
                dateEnd = df.iloc[k]['Fecha Fin']
                if dateCellS < dateEnd and not pass_x:
                    prevMax = k
                    dateStart = df.iloc[k]['Fecha Inicio']
                    numberofChapters = df.iloc[k]['Episodios'] - 1
                    daysBetween = (dateEnd - dateStart).days
                    if numberofChapters != 0:
                        daysBetween = daysBetween / numberofChapters
                    calculateDays(dateCellS, dateStart, daysBetween, dateCellE, i, prevMax)
                    pass_x = True
            if not pass_x:
                prevMax = i
                dateStart = dateCellS
                dateEnd = dateCellE
                numberofChapters = df.iloc[i]['Episodios'] - 1
                daysBetween = (dateEnd - dateStart).days
                if numberofChapters != 0:
                    daysBetween = daysBetween / numberofChapters
                episodeALoneList.append("")
        i = i + 1
    df['Cap. Sueltos'] = episodeALoneList


def calculateDays(dateCellS, dateStart, daysBetween, dateCellE, i, prevMax):
    minLongDate = (dateCellS - dateStart).days / daysBetween
    episodeALoneList.append(str(prevMax + 2) + ")" + str(int(minLongDate + 1)))

    numberofChapters = df.iloc[i]['Episodios']
    if numberofChapters != 1:
        daysBetween2 = int((dateCellE - dateCellS).days / (numberofChapters - 1))

        for p in range(2, numberofChapters + 1):
            dateCellS = dateCellS + datetime.timedelta(days=daysBetween2)
            minLongDate = (dateCellS - dateStart).days / daysBetween
            episodeALoneList[i] = episodeALoneList[i] + ", " + str(int(minLongDate + 1))


def excel(title):
    df.to_excel(title, index=False)

    wb = openpyxl.load_workbook(filename=title)
    worksheet = wb['Sheet1']
    worksheet.title = 'Lista Animes'

    for col in worksheet.columns:
        max_length = 0
        column = col[0].column_letter  # Get the column name
        if column == "M":
            continue
        k = 1
        for cell in col:
            if column == "A" and k != 1:
                worksheet[column + str(k)].hyperlink = worksheet["M" + str(k)].value
            if (column == "B" or column == "C" or column == "D") and k != 1:
                worksheet[column + str(k)].style = date_style
            if (column == "J" or column == "K") and k != 1:
                worksheet[column + str(k)].alignment = Alignment(wrapText=True)
            try:  # Necessary to avoid error on empty cells
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
            k = k + 1
        adjusted_width = max_length
        # adjusted_width = (max_length + 2) * 1.2
        worksheet.column_dimensions[column].width = adjusted_width
    worksheet.delete_cols(13, 1)
    worksheet.auto_filter.ref = worksheet.dimensions
    wb.save(title)
    os.system(title)


listTitle = list()
listTitleWithouthEspecialChars = list()
dateList = list()
numberdateFinishList = list()
episodeList = list()
durationList = list()
typeList = list()
studiosList = list()
urlList = list()
seenList = list()
opList = list()
endList = list()
episodeALoneList = list()
statusList = list()
base = "https://myanimelist.net"
baseTxt = "Txts\\"

characterNotPermited = getList("characterNotPermited.txt")
listProhibited = ["Adaptation:", "Adaptation (Light Novel)", "Adaptation (Manga)",
                  "Adaptation (Manhwa)", "Adaptation (Novel)", "Adaptation (One-shot)"]
listProhibited2 = ["Character:"]
nameSpecialProhibited = getList("nameSpecialProhibited.txt")

nameCategoryNotPermited = getList("AlternativeSetting.txt")
categoryNotPermited = ["Alternative Setting:"]

nameCategoryNotPermited2 = getList("AlternativeVersion.txt")
categoryNotPermited2 = ["Alternative Version:"]

nameCategoryNotPermited3 = getList("Other.txt")
categoryNotPermited3 = ["Other:"]

nameCategoryNotPermited4 = getList("ParentStory.txt")
categoryNotPermited4 = ["Parent Story:"]

nameProhibited = getList("nameProhibited.txt")
nameProhibitedIncluded = getList("nameProhibitedIncluded.txt")

baseAnime = "https://myanimelist.net/anime/"
fromUrl = [baseAnime + "31647/Ponkotsuland_Saga", baseAnime + "40034/Tenki_no_Ko_CMs",
           baseAnime + "33904/Suntory_Tennensui_CMs", baseAnime + "39511/Main_Actor",
           baseAnime + "5081/Bakemonogatari"]
toUrl = [baseAnime + "37521/Vinland_Saga", baseAnime + "38826/Tenki_no_Ko",
         baseAnime + "32281/Kimi_no_Na_wa", baseAnime + "37982/Domestic_na_Kanojo",
         baseAnime + "30514/Nisekoimonogatari"]
letterProhibited = ["\\", "/", ":", "*", "?", "<", ">", "|", '"']

idNumber = "31240/Re_Zero_kara_Hajimeru_Isekai_Seikatsu"
nameExcel = "Excel\\Ramas_Anime_" + datetime.date.today().strftime("%d_%m_%Y") + ".xlsx"
err = 1
date_style = NamedStyle(name='datetime', number_format='DD/MM/YYYY')
df = pd.DataFrame({})


def main():
    ramas(list(), datetime.date.today(), baseAnime + idNumber)
    getEps()
    excel(nameExcel)


main()