"""
Script Python

Goal Utama: - Dapat mengisi dataset acuan yang berisi list file excel yang diawasi ketepatan waktu dalam memperbarui filenya. []
            - Dapat menentukan file mana yang diperbarui dalam waktu yang diberikan dan juga mana yang terlambat []
            - Dapat diintegrasikan dengan Power BI []

Fungsi Pendukung: - Fungsi mendapatkan nama-nama file dan informasinya yang diawasi secara otomatis dari directory file tersebut disimpan []
                  - Fungsi memproses string nama file []
                  - Fungsi memproses datetime []
"""

# Import modul-modul penting
import pandas as pd
from openpyxl import load_workbook
import os


# Fungsi import dataset utama
def importEmptyMainDataset():
    return pd.read_excel("dataset/data_acuan.xlsx")

# Fungsi Mengisi dataset utama dengan informasi file-file yang diawasi
def fillEmptyMainDataset(mainDataset, listOfExcelFile):
    pass

# Fungsi untuk mendapatkan nama file excel
def getExcelFileName(excelFile):
    pass

# Fungsi traverse directory mulai dari parent untuk menemukan file dengan format '.xlsx'
def exploreDirectory():
    path = 'C:\Computer Science\PKL'
    listOfFiles = []
    for root, dirs, files in os.walk(path):
        for file in files:
            if file.endswith(".xlsx"):
                 listOfFiles.append(os.path.join(root, file))
    
    return listOfFiles
    
    # extension = 'xlsx'
    # os.chdir(path)
    # result = glob.glob('*.{}'.format(extension))
    # print(result)


# Fungsi Main
if __name__ == "__main__":
    mainDataset = importEmptyMainDataset()
    listOfExcelFile = exploreDirectory()
    