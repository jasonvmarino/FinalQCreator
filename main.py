import filereader
import os

file = os.listdir()
files = []
for item in file:
    if item[-3:] == 'txt':
        if item[-11:] != 'answers.txt':
            files.append(item[:-4])
for item in files:
    print(item)
    filereader.FileReader(item)
    file_list = [item + '.txt', item + ' answers.txt' + '.xlsx']
    inipath = os.getcwd() + chr(92)
    path = os.getcwd() + chr(92) + item + chr(92)
    access_rights = 0o755
    try:
        os.mkdir(path, access_rights)
    except:
        pass
    for items in file_list:
        os.rename(inipath + item, path + item)