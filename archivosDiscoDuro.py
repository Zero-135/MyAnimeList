import os
from pathlib import Path
from datetime import datetime
import openpyxl
import pandas as pd
import pyodbc
from moviepy.editor import VideoFileClip, ImageClip


def getlist(name):
    my_file = open(baseTxt + name, "r")
    data = my_file.read()
    data = data.replace("Â", "")
    my_file.close()
    return data.split("\n")


def error(name2, sub, more, empty):
    for l1 in listNoProcces:
        if name2 == l1:
            return True

    if not more:
        for l2 in listNoMyanimeList:
            if name2 == l2:
                return True

    if not empty:
        size = 0
        for ele in os.scandir(sub):
            size += os.path.getsize(ele)

        if size == 0:
            print(name2)
            return True

    return False


def getfiles(folder_path, foldername):
    global numberIdFile, numberIdFile, supported_extensions
    try:
        if not os.path.exists(roothBaseDownload + "\\" + foldername):
            os.makedirs(roothBaseDownload + "\\" + foldername)

        for ele in os.scandir(folder_path):
            name = ele.name
            path = ele.path

            if os.path.isdir(path):
                getfiles(path, foldername + "\\" + name)
            else:
                _, extension = os.path.splitext(path)

                if extension in supported_extensions:
                    query = ("insert into fileAnime values(" +
                             str(numberIdFile) + ",N'" +
                             name.replace("'", "''") + "','" +
                             str(numberIdSubFolder) + "',N'" +
                             path.replace("'", "''") + "',N'" +
                             foldername.replace("'", "''") +
                             "')"
                             )
                    numberIdFile = numberIdFile + 1
                    cursor.execute(query)

                    roothDownload = (roothBaseDownload + "\\" + foldername + "\\" + name + ".jpg")

                    if not os.path.isfile(roothDownload):
                        video = VideoFileClip(path)
                        video = video.resize(0.13)
                        frame = video.get_frame(video.duration / 2)
                        image = ImageClip(frame)
                        image.save_frame(roothDownload)

    except NameError:
        print(NameError)


def files(nameexcel, more, empty):
    global numberIdSubFolder, numberIdFile
    numberIdFolder = 1
    for path in paths:
        direc = sorted(Path(path).iterdir(), key=os.path.getctime, reverse=True)
        filtered = filter(lambda fichero: "[BD]" in os.path.basename(fichero), direc)
        listBD = list(filtered)
        for fichero in listBD:
            name = os.path.basename(fichero).replace("[BD]", "")
            date = datetime.fromtimestamp(os.path.getctime(fichero)).strftime('%Y-%m-%d %H:%M:%S')
            rootPathOriginal = os.path.abspath(fichero)
            rootPath = os.path.abspath(fichero).replace("'", "''")

            first = True
            subfolders = [f.path for f in os.scandir(fichero) if f.is_dir()]
            print(name)
            if len(subfolders) > 0:
                if os.path.basename(subfolders[0]) == "Anime":
                    subfolders = [f.path for f in os.scandir(subfolders[0]) if f.is_dir()]

                for sub in subfolders:
                    name2 = os.path.basename(sub)
                    roothOriginal = os.path.abspath(sub)
                    rootPathSub = os.path.abspath(sub).replace("'", "''")

                    index = name2.find(".-")

                    if index > -1:
                        name2 = name2[index + 2::]

                    if error(name2, sub, more, empty):
                        continue

                    if first:
                        if 'Yuru' in name:
                            name = name
                        listName.append(str(numberIdFolder) + ".-" + name)
                        listDate.append(date)
                        first = False
                        query = ("insert into Folders values(" +
                                 str(numberIdFolder) + ",N'" +
                                 name.replace("'", "''") + "',N'" +
                                 rootPath +
                                 "')"
                                 )
                        cursor.execute(query)
                    else:
                        listName.append("")
                        listDate.append("")

                    listSub.append(name2)
                    query = ("insert into SubFolders values(" +
                             str(numberIdSubFolder) + ",N'" +
                             name2.replace("'", "''") + "','" +
                             str(numberIdFolder) + "',N'" +
                             rootPathSub +
                             "')"
                             )
                    cursor.execute(query)
                    getfiles(roothOriginal, name + "\\" + name2)
                    numberIdSubFolder = numberIdSubFolder + 1
            else:
                listName.append(str(numberIdFolder) + ".-" + name)
                listDate.append(date)
                query = ("insert into Folders values(" +
                         str(numberIdFolder) + ",N'" +
                         name.replace("'", "''") + "','" +
                         rootPath +
                         "')"
                         )
                cursor.execute(query)

                listSub.append(name)
                query = ("insert into SubFolders values(" +
                         str(numberIdSubFolder) + ",N'" +
                         name.replace("'", "''") + "','" +
                         str(numberIdFolder) + "','" +
                         rootPath +
                         "')"
                         )
                cursor.execute(query)
                getfiles(rootPathOriginal, name)
                numberIdSubFolder = numberIdSubFolder + 1
            numberIdFolder = numberIdFolder + 1

    conn.commit()
    df = pd.DataFrame({"Nombre": listName, "SubDirectorio": listSub, "Fecha": listDate})
    df.to_excel(nameexcel, index=False)

    wb = openpyxl.load_workbook(filename=nameexcel)
    worksheet = wb.active

    for col in worksheet.columns:
        max_length = 0
        column = col[0].column_letter  # Get the column name
        for cell in col:
            try:  # Necessary to avoid error on empty cells
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except NameError:
                pass
        adjusted_width = (max_length + 2) * 1.2
        worksheet.column_dimensions[column].width = adjusted_width

    wb.save(nameexcel)
    os.system(nameexcel)


baseTxt = "Txts\\"
listNoProcces = getlist("listNoProcces.txt")
listNoMyanimeList = getlist("listNoMyanimeList.txt")
paths = ["I:/", "H:/", "G:/"]
listName = list()
listSub = list()
listDate = list()
server = '(LocalDb)\\MSSQLLocalDB'
bd = 'animeList'
numberIdSubFolder = 1
numberIdFile = 1
# supported_extensions = ['.mkv', '.mp4', '.avi', '.mpg', '.mpeg']
supported_extensions = ['.mov', '.mp4', '.mpg', '.mpeg', '.flv', '.wmv', '.mkv', '.avi']
# roothBaseDownload = r"C:\Users\Walter Rivas\source\repos\AplicacionAnime\AplicacionAnime\Images"
roothBaseDownload = r"C:\AppAnime\AppAnime\Images"
try:
    conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=' + server + ';DATABASE=' + bd +
                          ';Integrated Security=True')
    cursor = conn.cursor()
    cursor.execute("truncate table Folders")
    cursor.execute("truncate table SubFolders")
    cursor.execute("truncate table fileAnime")
    conn.commit()
except NameError:
    print(NameError)

files("Excel\\Nombres_Carpetas.xlsx", True, True)

# if "AVC" in codec_id:
#         extractName = extractName + ".264"
#     elif "HEVC" in codec_id:
#         extractName = extractName + ".hevc"
#     elif "V_VP8" in codec_id:
#         extractName = extractName + ".ivf"
#     elif "V_VP9" in codec_id:
#         extractName = extractName + ".ivf"
#     elif "V_AV1" in codec_id:
#         extractName = extractName + ".ivf"
#     elif "V_MPEG1" in codec_id:
#         extractName = extractName + ".mpg"
#     elif "V_MPEG2" in codec_id:
#         extractName = extractName + ".mpg"
#     elif "V_REAL" in codec_id:
#         extractName = extractName + ".rm"
#     elif "V_THEORA" in codec_id:
#         extractName = extractName + ".ogg"
#     elif "V_MS/VFW/FOURCC" in codec_id:
#         extractName = extractName + ".avi"
#     elif "AAC" in codec_id:
#         extractName = extractName + ".aac"
#     elif "A_AC3" in codec_id:
#         extractName = extractName + ".ac3"
#     elif "A_EAC3" in codec_id:
#         extractName = extractName + ".eac3"
#     elif "ALAC" in codec_id:
#         extractName = extractName + ".caf"
#     elif "DTS" in codec_id:
#         extractName = extractName + ".dts"
#     elif "FLAC" in codec_id:
#         extractName = extractName + ".flac"
#     elif "MPEG/L2" in codec_id:
#         extractName = extractName + ".mp2"
#     elif "MPEG/L3" in codec_id:
#         extractName = extractName + ".mp3"
#     elif "OPUS" in codec_id:
#         extractName = extractName + ".ogg"
#     elif "PCM" in codec_id:
#         extractName = extractName + ".wav"
#     elif "REAL" in codec_id:
#         extractName = extractName + ".ra"
#     elif "TRUEHD" in codec_id:
#         extractName = extractName + ".thd"
#     elif "MLP" in codec_id:
#         extractName = extractName + ".mlp"
#     elif "TTA1" in codec_id:
#         extractName = extractName + ".tta"
#     elif "VORBIS" in codec_id:
#         extractName = extractName + ".ogg"
#     elif "WAVPACK4" in codec_id:
#         extractName = extractName + ".wv"
#     elif "PGS" in codec_id:
#         extractName = extractName + ".sup"
#     elif "ASS" in codec_id:
#         extractName = extractName + ".ass"
#     elif "SSA" in codec_id:
#         extractName = extractName + ".ssa"
#     elif "UTF8" in codec_id:
#         extractName = extractName + ".srt"
#     elif "ASCII" in codec_id:
#         extractName = extractName + ".srt"
#     elif "VOBSUB" in codec_id:
#         extractName = extractName + ".sub"
#     elif "S_KATE" in codec_id:
#         extractName = extractName + ".ogg"
#     elif "USF" in codec_id:
#         extractName = extractName + ".usf"
#     elif "WEBVTT" in codec_id:
#         extractName = extractName + ".vtt"
