import openpyxl
import pandas as pd
from openpyxl import Workbook,load_workbook
from openpyxl.utils import get_column_letter
import tkinter as tk
from tkinter import filedialog
import os
root = tk.Tk()
root.withdraw()

print("Select Synergy 'Google Photos' export file")
file = filedialog.askopenfilename()
print("File:",file)
folderName = input("Enter name of desktop folder to look in? ")


def addColumnHeader(file):
    wb = load_workbook(file) #Load Workbook
    ws = wb.active #Worksheet
    ws['D1'].value = "File_Location"
    wb.save(file)

# Concatenates the ID field with the file location.
def populateFileLocation(file,folderName):
    wb = load_workbook(file) #Load Workbook
    ws = wb.active #Worksheet
    for row in range (2,ws.max_row+1): 
        for col in range (4,5): # column D
            char = get_column_letter(col)
            id =  ws[get_column_letter(3) + str(row)].value
            ws[char + str(row)].value = r'C:\Users\zechaaron\OneDrive - Osseo Area Schools\Desktop\{}\{}.jpg'.format(folderName,id)
    wb.save(file)

#Export File as CSV
def convertToCSV(file):
    fileData = pd.read_excel(file,sheet_name='QRY801')
    fileData.to_csv("GooglePhotos_GAM.csv",index=False)
    os.remove(file) #delete XLSX file 


#Main Program 

if __name__ == "__main__":
    print("Starting...")
    addColumnHeader(file)
    populateFileLocation(file,"photos")
    convertToCSV(file)
    print("finished")
