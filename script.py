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
import datetime

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

# Fungsi import dataset utama yang telah diisi
def importFilledMainDataset():
    return pd.read_excel("output.xlsx")

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

# Fungsi utama untuk menemukan file paling terupdate untuk kasus file yang diupdate dengan cara "append new file"
# Mengembalikan True jika file yang sedang diperiksa adalah file terupdate dari pada file yang sudah terdata di file acuan
def compareFilesDatetime(newFileDate, latestFileDate):
    return newFileDate > latestFileDate

# Fungsi untuk membandingkan tahun dan tanggal file dengan waktu sekarang. True jika telat
def compareMonthDay(FileTargetMonth, fileTargetDay, fileRealizationTime):
    currentDate = utcToLocal(datetime.datetime.now())
    fileRealizationTime = datetime.datetime.strptime(fileRealizationTime, "%Y-%m-%d %H:%M:%S")
    if currentDate.year == fileRealizationTime.year:
        if fileRealizationTime.month == FileTargetMonth:
            if fileRealizationTime.day > fileTargetDay:
                return True
        elif fileRealizationTime.month > FileTargetMonth:
            return True
    return False
                
# Fungsi membandingkan tanggal. True jika berada di tanggal yang sama
def compareDate(fileRealizationTime):
    currentDate = utcToLocal(datetime.datetime.now())
    return currentDate.day == fileRealizationTime.day

# Fungsi membandingkan jam dan menit
def compareHour(fileTime, fileRealizationTime):
    fileRealizationTime = datetime.datetime.strptime(fileRealizationTime, "%Y-%m-%d %H:%M:%S")
    if compareDate(fileRealizationTime):
        if fileTime.hour < fileRealizationTime.hour:
            return True
    return False

# ====================================================== Fungsi-Fungsi Utama ======================================================

# Fungsi Mengisi dataset utama dengan informasi file-file excel
def fillEmptyMainDataset(mainDataset, listOfExcelFile):
    for i in listOfExcelFile:
        wb = load_workbook(i)
        fileName = getExcelFileName(i)
        fileNameFormatted = formattingFileName(fileName)
        if fileNameFormatted not in mainDataset.File_Name.values:
            mainDataset.loc[len(mainDataset.index)] = [fileNameFormatted, fileName , i, 'Update Existing File', wb.properties.lastModifiedBy, '', '', '', utcToLocal(wb.properties.modified).strftime("%Y-%m-%d %H:%M:%S"), '']
        else:
            mainDataset.at[len(mainDataset.index)-1, 'Modification_Type'] = 'Update By Adding A New File'
            lastModifiedNewFile = utcToLocal(wb.properties.modified).strftime("%Y-%m-%d %H:%M:%S")
            if compareFilesDatetime(lastModifiedNewFile, mainDataset.iloc[len(mainDataset.index)-1]['Realisasi']):
                mainDataset.loc[len(mainDataset.index)-1] = [fileNameFormatted, fileName , i, 'Update By Adding A New File', wb.properties.lastModifiedBy, '', '', '', lastModifiedNewFile, '']
    return mainDataset

# Fungsi proses pengisian kolom SLA (Kategoriasi)
def slaCategorizationProcess(rowData):
    if rowData['Update_Periode'] == "Daily":
        fileTargetTime = re.findall('\d{2}:\d{2}', rowData['Target_Update'])
        fileTargetTime = datetime.datetime.strptime(fileTargetTime[0], '%H:%M')
        if compareHour(fileTargetTime, rowData['Realisasi']):
            return "Miss"
        else:
            return "Met"
    elif rowData['Update_Periode'] == "Monthly":
        pass
    elif rowData['Update_Periode'] == "Yearly":
        fileTargetMonth = monthNum(re.findall(reMonthName(), rowData['Target_Update'])[0])
        fileTargetDay = int(re.findall("[0-2][0-9]|[0-9]", rowData['Target_Update'])[0])
        if compareMonthDay(fileTargetMonth, fileTargetDay, rowData['Realisasi']):
            return "Miss"
        else:
            return "Met"

# Fungsi utama pengisian kolom SLA (Kategoriasi)
def slaCategorization(mainDataset):
    for index, rowData in mainDataset.iterrows():
        mainDataset.at[index, 'SLA_(Met/Miss)'] = slaCategorizationProcess(rowData)
    return mainDataset


# ====================================================== Fungsi Main ======================================================

# Fungsi Main
if __name__ == "__main__":

    #Proses mengisi data target 
    # mainDataset = importEmptyMainDataset()
    # listOfExcelFile = exploreDirectory()
    # mainDataset = fillEmptyMainDataset(mainDataset, listOfExcelFile)
    # mainDataset.to_excel('output.xlsx', index=False)

    # Proses kategorisasi SLA
    mainDataset = importFilledMainDataset()
    mainDataset = mainDataset.astype(str)
    mainDataset = slaCategorization(mainDataset)
    mainDataset.to_excel('output2.xlsx', index=False)
    
    
    