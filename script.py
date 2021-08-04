"""
Script Python

Fungsi Utama: - Dapat mengisi dataset acuan yang berisi list file excel yang diawasi ketepatan waktu dalam memperbarui filenya. [X]
              - Dapat menentukan file mana yang diperbarui dalam waktu yang diberikan dan juga mana yang terlambat []
              - Dapat diintegrasikan dengan Power BI []

Fungsi Pendukung: - Fungsi eksplorasi direktori untuk mendapatkan nama-nama file excel dan propertynya secara otomatis dari directory file tersebut disimpan [X]
                  - Fungsi memproses string nama file untuk file yang cara updatenya adalah menambahkan file baru [X]
                  - Fungsi memproses datetime untuk pencocokan target waktu dan realisasi []
"""

# ====================================================== Import Library ======================================================

import pandas as pd
from openpyxl import load_workbook
import os
import pytz
import re

# ====================================================== Variabel Global ======================================================

monthDict = {'Januari': 1, 'Februari': 2, 'Maret': 3, 'April': 4, 'Mei': 5, 'Juni': 6, 'Juli': 7,
                'Agustus': 8, 'September': 9, 'Oktober': 10, 'November': 11, 'Desember': 12}

#monthList = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember']

# ====================================================== Fungsi-Fungsi Pendukung ======================================================

# Fungsi mengubah datetime utc menjadi local time
def utcToLocal(time):
    localTz = pytz.timezone('Asia/Jakarta')
    localDt = time.replace(tzinfo=pytz.utc).astimezone(localTz)
    return localTz.normalize(localDt)

# ====================================================== Fungsi-Fungsi Input File-File Excel yang Terlibat ======================================================

# Fungsi import dataset utama
def importEmptyMainDataset():
    return pd.read_excel("Pemeriksa Ketepatan Waktu Update\dataset\data_acuan.xlsx")

# Fungsi untuk mendapatkan nama file excel yang sudah dipisah dengan pathnya
def getExcelFileName(excelFile):
    return [item for item in excelFile.split("\\")][-1]
     
# Fungsi traverse directory mulai dari parent untuk menemukan file dengan format '.xlsx'
def exploreDirectory():
    path = 'C:\Computer Science\PKL'
    listOfFile = []
    for root, dirs, files in os.walk(path):
        for file in files:
            if file.endswith(".xlsx"):
                 listOfFile.append(os.path.join(root, file))

    return listOfFile


# ====================================================== Fungsi-Fungsi Memproses String Nama File ======================================================

# Fungsi konversi nama bulan ke bentuk numerik (Januari = 1, Februari = 2, etc)
def monthNum(month):
    return monthDict[month]

# Fungsi yang memberikan regex untuk mencari bulan pada suatu string
def reMonthName():
    return "Januari|Februari|Maret|April|Mei|Juni|Juli|Agustus|September|Oktober|November|Desember"

# Fungsi yang memberikan regex untuk mencari tahun 4 digit pada suatu string
def reYear4Digits():
    return "[1-3][0-9]{3}"

# Fungsi yang memberikan regex untuk mencari bulan dan tahun berbentuk angka pada suatu string 
def reMonthAndYear():
    return "[0-1][0-9][1-3][0-9]"

# Fungsi untuk mengubah tahun dari nama file ke bentuk umumnya (year) 
def formattingYear4Digits(fileName, hasYear4Digits):
    return fileName.replace(hasYear4Digits[0], 'year')

# Fungsi untuk mengubah nama bulan dari nama file ke bentuk umumnya (month) 
def formattingMonthName(fileName, hasMonthName):
    return fileName.replace(hasMonthName[0], 'month')

# Fungsi untuk mengubah bulan dan tahun dari nama file ke bentuk umumnya (mmyy) 
def formattingMonthAndYear(fileName, hasMonthAndYear):
    return fileName.replace(hasMonthAndYear[0], 'mmyy')

# Fungsi untuk generalisasi nama file yang memiliki isi yang sama. Misal file dengan nama "Laporan Keuangan 0221", akan menjadi "Laporan Keuangan mmyy" dimana
# mm adalah bulan dan yy adalah tahun
def formattingFileName(fileName):
    hasMonthName = re.findall(reMonthName(), fileName)
    hasYear = re.findall(reYear4Digits(), fileName)
    hasMonthAndYear = re.findall(reMonthAndYear(), fileName)

    if hasMonthName:
        fileName =  formattingMonthName(fileName, hasMonthName)
    if hasYear:
        fileName = formattingYear4Digits(fileName, hasYear)
    elif hasMonthAndYear:
        fileName = formattingMonthAndYear(fileName, hasMonthAndYear)
    return fileName

# ====================================================== Fungsi-Fungsi Datetime ======================================================

#Fungsi untuk menemukan file paling terupdate pada bagian bulan untuk kasus file yang diupdate dengan cara "append new file"
def compareMonth(newFileName, latestFileName):
    pass

#Fungsi untuk menemukan file paling terupdate pada bagian tahun untuk kasus file yang diupdate dengan cara "append new file"
def compareYear(newFileName, latestFileName):
    pass

# ====================================================== Fungsi-Fungsi Utama ======================================================

# Fungsi Mengisi dataset utama dengan informasi file-file excel
def fillEmptyMainDataset(mainDataset, listOfExcelFile):
    latestFileNum = 0
    for i in listOfExcelFile:
        wb = load_workbook(i)
        fileName = getExcelFileName(i)
        fileNameFormatted = formattingFileName(fileName)
        if fileNameFormatted not in mainDataset.File_Name.values:
            mainDataset.loc[len(mainDataset.index)] = [fileNameFormatted, fileName , i, 'Update Existing File', wb.properties.lastModifiedBy, '', '', '', utcToLocal(wb.properties.modified).strftime("%Y-%m-%d %H:%M:%S"), '']
        else:
            mainDataset.at[len(mainDataset.index)-1, 'Modification_Type'] = 'Update By Adding A New File'

    return mainDataset

# ====================================================== Fungsi Main ======================================================

# Fungsi Main
if __name__ == "__main__":
    mainDataset = importEmptyMainDataset()
    listOfExcelFile = exploreDirectory()
    mainDataset = fillEmptyMainDataset(mainDataset, listOfExcelFile)
    mainDataset.to_excel('output.xlsx', index=False)
    