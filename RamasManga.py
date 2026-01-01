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
        name = (soup.find_all('span', class_='h1-title')[0].text).replace("‚òÖ", "★")
    except:
        print("Error:" + str(err) + " " + link)
        if link == "/manga//":
            return
        err = err + 1
        time.sleep(5 * 60)
        ramas(listaStatus, date_x, link)
        return

    urlList.append(url)
    listTitle.append(name)
    seenList.append(date_x)


    statusName = ""
    name2 = name
    for j in letterProhibited:
        name2 = name2.replace(j, "")

    listTitleWithouthEspecialChars.append(name2)

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
            if t.text == "Volumes:":
                volumeList.append(k.text.rstrip().lstrip())
            if t.text == "Chapters:":
                episodesString = k.text.rstrip().lstrip()
                if episodesString != "Unknown":
                    episodeList.append(int(episodesString))
                else:
                    episodeList.append(999)
            if t.text == "Published:":
                textdate = k.text.replace(",", "").replace("  ", " ")
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
    df = pd.DataFrame({
        'Nombre': listTitleWithouthEspecialChars,
        'Fecha Visto': seenList,
        'Fecha Inicio': dateList,
        'Fecha Fin': numberdateFinishList,
        'Tipo': typeList,
        'Episodios': episodeList,
        'Volumenes': volumeList,
        'Estado': statusList,
        'Url': urlList,

    })
    df.sort_values(by=['Fecha Visto', 'Fecha Inicio', 'Nombre'], ascending=[False, True, False], inplace=True)


def excel(title):
    df.to_excel(title, index=False)

    wb = openpyxl.load_workbook(filename=title)
    worksheet = wb['Sheet1']
    worksheet.title = 'Lista Animes'

    for col in worksheet.columns:
        max_length = 0
        column = col[0].column_letter  # Get the column name
        if column == "I":
            continue
        k = 1
        for cell in col:
            if column == "A" and k != 1:
                worksheet[column + str(k)].hyperlink = worksheet["I" + str(k)].value
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
    worksheet.delete_cols(9, 1)
    worksheet.auto_filter.ref = worksheet.dimensions
    wb.save(title)
    os.system(title)


listTitle = list()
listTitleWithouthEspecialChars = list()
dateList = list()
numberdateFinishList = list()
episodeList = list()
volumeList = list()
typeList = list()
urlList = list()
seenList = list()
episodeALoneList = list()
statusList = list()
base = "https://myanimelist.net"
baseTxt = "Txts\\"

characterPermited = getList("characterNotPermited.txt")
listProhibited = ["Adaptation:"]
listProhibited2 = ["Character:"]

nameCategoryNotPermited = getList("AlternativeSetting.txt")
categoryNotPermited = ["Alternative setting:"]

nameCategoryNotPermited2 = getList("AlternativeVersion.txt")
categoryNotPermited2 = ["Alternative version:"]

nameCategoryNotPermited3 = getList("Other.txt")
categoryNotPermited3 = ["Other:"]

nameProhibited = getList("nameProhibited.txt")
nameProhibitedIncluded = getList("nameProhibitedIncluded.txt")

baseManga = "/manga/"
fromUrl = [baseManga + "31647/Ponkotsuland_Saga", baseManga + "40034/Tenki_no_Ko_CMs",
           baseManga + "33904/Suntory_Tennensui_CMs"]
toUrl = [baseManga + "37521/Vinland_Saga", baseManga + "38826/Tenki_no_Ko", baseManga + "32281/Kimi_no_Na_wa"]
letterProhibited = ["\\", "/", ":", "*", "?", "<", ">", "|", '"']

idNumber = "1264/Higurashi_no_Naku_Koro_ni__Onisarashi-hen"
nameExcel = "Excel\\Ramas_Manga_" + datetime.date.today().strftime("%d_%m_%Y") + ".xlsx"
err = 1
date_style = NamedStyle(name='datetime', number_format='DD/MM/YYYY')
df = pd.DataFrame({})

ramas(list(), datetime.date.today(), baseManga + idNumber)
getEps()
excel(nameExcel)
