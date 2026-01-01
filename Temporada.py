import os
import re

import openpyxl
import pandas as pd
import requests
from bs4 import BeautifulSoup


def urllink(url):
    try:
        return requests.get(url)
    except:
        return urllink(url)


def downloadImage(url, name, name2):
    nameFolder = year + "-" + temp
    if not os.path.isdir(nameFolder):
        os.mkdir(nameFolder)

    nameType = nameFolder + "/" + name2
    if not os.path.isdir(nameType):
        os.mkdir(nameType)

    nameFile = nameType + "/" + name + ".png"
    if not os.path.exists(nameFile):
        response = urllink(url)
        file = open(nameFile, "wb")
        file.write(response.content)
        file.close()


def Temporada():
    url = 'https://myanimelist.net/anime/season/' + year + "/" + temp
    page = urllink(url)
    soup = BeautifulSoup(page.content, 'html.parser')

    divGeneral = soup.find_all("div", class_=re.compile("^seasonal-anime-list js-seasonal-anime-list"))

    for divs in divGeneral:
        namesTitle = divs.find_all('div', class_='anime-header')
        namesAnime = divs.find_all('a', class_='link-title')
        imageLink = divs.find_all('img')
        listTitle1 = []
        k = 0
        for n in namesAnime:
            name = n.text
            for j in letterProhibited:
                name = name.replace(j, "")
            linkImage = imageLink[k].get('src')
            if linkImage is None:
                linkImage = imageLink[k].get('data-src')
                downloadImage(linkImage, name, namesTitle[0].text)
            listTitle1.append(name)
            k = k + 1

        listTitle1.sort()

        for lt in listTitle1:
            linkFolder.append("..\\" + year + "-" + temp + "\\" + namesTitle[0].text + "\\" + lt + ".png")
            listType.append(namesTitle[0].text)
            listTitle.append(lt)


def Excel():
    df = pd.DataFrame({
        "Tipo": listType,
        "Nombre": listTitle,
        "Url": linkFolder
    })
    df.to_excel(titleExcel, index=False)
    wb = openpyxl.load_workbook(filename=titleExcel)
    worksheet = wb.active

    for col in worksheet.columns:
        max_length = 0
        column = col[0].column_letter  # Get the column name
        k = 1
        for cell in col:
            if column == "B" and k != 1:
                worksheet[column + str(k)].hyperlink = worksheet["C" + str(k)].value
            try:  # Necessary to avoid error on empty cells
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
            k = k + 1
        adjusted_width = (max_length + 2) * 1.2
        worksheet.column_dimensions[column].width = adjusted_width

    worksheet.delete_cols(3, 1)
    wb.save(titleExcel)
    os.system(titleExcel)


temp = "spring"
year = "2022"
titleExcel = "Excel\\Temporada.xlsx"
letterProhibited = ["\\", "/", ":", "*", "?", "<", ">", "|"]
listTitle1 = list()
listTitle = list()
listType = list()
linkFolder = list()
Temporada()
Excel()
