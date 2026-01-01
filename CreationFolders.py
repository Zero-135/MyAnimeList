import os
import time


def getList(name):
    my_file = open(baseTxt + name, "r")
    data = my_file.read()
    my_file.close()
    return data.split("\n")


baseTxt = "Txts\\"
foldersList = getList("NameDirectories.txt")
foldersList.reverse()
parent_dir = "C:/Users/Walter Rivas/Music/Audio-Anime/"

for directory in foldersList:
    path = os.path.join(parent_dir, directory)
    os.mkdir(path)
    time.sleep(2)
