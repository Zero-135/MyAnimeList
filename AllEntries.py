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
    my_file = open(baseTxt + name, "r")
    data = my_file.read()
    my_file.close()
    return data.split("\n")


def urllink(url):
    try:
        return requests.get(url)
    except:
        return urllink(url)


def convertDate(timestr):
    if timestr == "Not available":
        return datetime.date(year=2099, day=1, month=1)
    else:
        if timestr == "?":
            return datetime.date.today()

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
    url = base + link
    page = urllink(url)
    soup = BeautifulSoup(page.content, 'html.parser')
    try:
        spanTags = soup.find_all('span', class_='title-english')
        for span in spanTags:
            span.decompose()

        nameAnime = ""
        nameManga = ""

        allNameAnime = soup.find_all('h1', class_='title-name h1_bold_none')
        if len(allNameAnime) > 0:
            nameAnime = (allNameAnime[0].text).replace("‚òÖ", "★")

        allNameManga = soup.find_all('span', class_='h1-title')
        if len(allNameManga) > 0:
            nameManga = allNameManga[0].text

    except:
        print("Error:" + str(err) + " " + link)
        if link == "/anime//" or link == "/manga//":
            return
        err = err + 1
        time.sleep(5 * 60)
        ramas(listaStatus, date_x, link)
        return

    urlList.append(url)
    seenList.append(date_x)

    opnening = soup.find_all('div', class_='opnening')
    if len(opnening):
        allTr = opnening[0].find_all('table')[1].find_all('tr')
        opname = ""
        for tr in allTr:
            allTd = tr.find_all('td')
            if len(allTd) > 1:
                if opname != "":
                    opname = opname + "\n"
                opname = opname + allTd[1].extract().text.rstrip().lstrip()
        opList.append(opname)
    else:
        opList.append("")

    ending = soup.find_all('div', class_='ending')
    if len(ending):
        allTr = ending[0].find_all('table')[0].find_all('tr')
        endname = ""
        for tr in allTr:
            allTd = tr.find_all('td')
            if len(allTd) > 1:
                if endname != "":
                    endname = endname + "\n"
                endname = endname + allTd[1].extract().text.rstrip().lstrip()
        endList.append(endname)
    else:
        endList.append("")

    listTd = soup.find_all('td', class_='borderClass')
    h2List = listTd[0].find_all('h2')
    i = 0
    h2 = []
    for j in h2List:
        if j.text == "Information":
            h2 = h2List[i]
        i = i + 1

    spans = h2.find_next_siblings()
    name = ""
    for k in spans:
        for t in k.select('span'):
            t.extract()
            if t.text == "Type:":
                if k.text.rstrip().lstrip() == "Manga" or k.text.rstrip().lstrip() == "Light Novel":
                    name = nameManga
                else:
                    name = nameAnime
                typeList.append(k.text.rstrip().lstrip())
            if t.text == "Studios:":
                studiosList.append(k.text.rstrip().lstrip())
            if t.text == "Duration:":
                durationList.append(k.text.rstrip().lstrip())
            if t.text == "Volumes:":
                volumeString = k.text.rstrip().lstrip()
                if volumeString != "Unknown":
                    volumeList.append(int(volumeString))
                else:
                    volumeList.append(999)
            if t.text == "Episodes:":
                episodesString = k.text.rstrip().lstrip()
                if episodesString != "Unknown":
                    episodeList.append(int(episodesString))
                else:
                    episodeList.append(999)
            if t.text == "Chapters:":
                chaptersString = k.text.rstrip().lstrip()
                if chaptersString != "Unknown":
                    chapterList.append(int(chaptersString))
                else:
                    chapterList.append(999)
            if t.text == "Published:":
                textdate = k.text.replace(",", "").replace("  ", " ")
                arrDate = textdate.split("to")
                textdate = arrDate[0].rstrip().lstrip()
                numberdate = convertDate(textdate)
                if len(arrDate) > 1:
                    numberdateFinish = convertDate(arrDate[1].rstrip().lstrip())
                else:
                    numberdateFinish = numberdate
                # publisheddateList.append(numberdate)
                # publishedFinishList.append(numberdateFinish)
                dateList.append(numberdate)
                numberdateFinishList.append(numberdateFinish)
            if t.text == "Aired:":
                textdate = k.text.replace(",", "")
                arrDate = textdate.split("to")
                textdate = arrDate[0].rstrip().lstrip()
                numberdate = convertDate(textdate)
                if len(arrDate) > 1:
                    numberdateFinish = convertDate(arrDate[1].rstrip().lstrip())
                else:
                    numberdateFinish = numberdate
                dateList.append(numberdate)
                numberdateFinishList.append(numberdateFinish)

    if link in fromUrl:
        link = toUrl[fromUrl.index(link)]
        ramas(listaStatus, date_x, link)

    if name == nameManga:
        studiosList.append("")
        episodeList.append("")
        durationList.append("")
        #dateList.append("")
        #numberdateFinishList.append("")
    else:
        volumeList.append("")
        chapterList.append("")
        #publisheddateList.append("")
        #publishedFinishList.append("")

    listTitle.append(name)
    name2 = name
    for j in letterProhibited:
        name2 = name2.replace(j, "")

    listTitleWithouthEspecialChars.append(name2)

    statusName = ""
    if len(listaStatus) > 0:
        try:
            index = listaStatus[0].index(name2)
            statusName = listaStatus[1][index]
        except:
            pass
    statusList.append(statusName)

    if name in nameProhibitedIncluded:
        return

    tableGeneral = soup.find_all("table", class_=re.compile("anime_detail_related_anime"))
    for table in tableGeneral:
        tdList = table.find_all('td', class_='fw-n')
        for td in tdList:
            if td.text == "Character:":
                print(name)

            # if not "Piece" in name:
            #    print(name)

            if (td.text in listProhibited) or \
                    (td.text in listProhibited2 and name not in characterPermited) or \
                    (td.text in categoryNotPermited and name in nameCategoryNotPermited) or \
                    (td.text in categoryNotPermited2 and name in nameCategoryNotPermited2) or \
                    (td.text in categoryNotPermited3 and name in nameCategoryNotPermited3):
                continue
            else:
                alink = td.find_next_siblings("td")[0].find_all("a", href=True)
                for a in alink:
                    finalName = (a.text).replace("‚òÖ", "★")
                    linkIntro = base + a['href']
                    if not (linkIntro in urlList) and not (finalName in nameProhibited):
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
        'Volumenes': volumeList,
        'Capitulos':chapterList,
        #'Fecha Incio Manga/Novela': publisheddateList,
        #'Fecha Fin Manga/Novela': publishedFinishList,
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

        if df.iloc[i]['Tipo'] == "Manga" or df.iloc[i]['Tipo'] == "Light Novel":
            episodeALoneList.append("")
            continue

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
        if column == "O":
            continue
        k = 1
        for cell in col:
            if column == "A" and k != 1:
                worksheet[column + str(k)].hyperlink = worksheet["O" + str(k)].value
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
    worksheet.delete_cols(15, 1)
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
volumeList = list()
chapterList = list()
publisheddateList = list()
publishedFinishList = list()
base = "https://myanimelist.net"
baseTxt = "Txts\\"

characterPermited = getList("characterNotPermited.txt")
#listProhibited = ["Adaptation:"]
listProhibited = []
listProhibited2 = ["Character:"]

nameCategoryNotPermited = getList("AlternativeSetting.txt")
categoryNotPermited = ["Alternative setting:"]

nameCategoryNotPermited2 = getList("AlternativeVersion.txt")
categoryNotPermited2 = ["Alternative version:"]

nameCategoryNotPermited3 = getList("Other.txt")
categoryNotPermited3 = ["Other:"]

nameProhibited = getList("nameProhibited.txt")
nameProhibitedIncluded = getList("nameProhibitedIncluded.txt")

idNumber = "16498/Shingeki_no_Kyojin"
baseAnime = "/anime/"
fromUrl = [baseAnime + "31647/Ponkotsuland_Saga", baseAnime + "40034/Tenki_no_Ko_CMs",
           baseAnime + "33904/Suntory_Tennensui_CMs"]
toUrl = [baseAnime + "37521/Vinland_Saga", baseAnime + "38826/Tenki_no_Ko", baseAnime + "32281/Kimi_no_Na_wa"]
letterProhibited = ["\\", "/", ":", "*", "?", "<", ">", "|", '"']

nameExcel = "Excel\\Ramas_Union_" + datetime.date.today().strftime("%d_%m_%Y") + ".xlsx"
err = 1
date_style = NamedStyle(name='datetime', number_format='DD/MM/YYYY')
df = pd.DataFrame({})

ramas(list(), datetime.date.today(), baseAnime + idNumber)
getEps()
excel(nameExcel)
