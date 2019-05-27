''' -------------------------------- Last Modified on Mon, 24 Mai 2019  --------------------------------'''
from typing import List, Any

''' -------------------------------- Author: Markos Gavalas             --------------------------------'''

'''The following scipt is parsing particular informations from .xlsx files and put them into .csv files'''
'''It is important to close the Excel files that you want to parse before you run the program (otherwise it will give an error)'''

from openpyxl import load_workbook
import csv
import os
import numpy as np

try:
    from StringIO import StringIO
except ImportError:
    from io import StringIO

############################################################## configs

config = {
    "Zufriedenheit mit dem Arbeitsplatz" : {
        "schema": ["Zahlenwert von 1-10", "Anmerkungen"],
        "sheet": "Mitarbeitergespräch",
        "inter_rows": 176
    },
       "Zufriedenheit mit der Betreuung durch den Vorgesetzten (TL/AL)" : {
        "schema": ["Zahlenwert von 1-10", "Anmerkungen"],
        "sheet": "Mitarbeitergespräch",
        "inter_rows": 186
    }
}


############################################################### global variables

# Please insert the directory where the .xlsx files are in
# if nothing is given the program will search for the files that are in the directory of the script
#xslx_dir = "C:/Users/gavam/Desktop/use_case"
xslx_dir = "."
# Please insert the directory where the .csv files will be created
# if nothing is given the program will search for the files that are in the directory of the script
csv_dir = "C:/Users/gavam/Desktop/use_case"

# Please insert here the Column name output you would like to have
output_cols = ["Zahlenwert von 1-10", "Anmerkungen"]

# Please insert here the excel sheet rows that correspond to the information you want to parse:

#################################################################################
xslxs = [f for f in os.listdir(xslx_dir) if f.endswith("xlsx")]
print (xslxs)
#Table with all the info
table = []

#Function
def parseri(xslx, SheetName, minmaxr):
        wb = load_workbook(xslx)
        sheet = wb[SheetName]
        lista = []
        for row in sheet.iter_rows(min_row=minmaxr, max_row=minmaxr, min_col=40, max_col=45):
            for cell in row:
                if cell.value is not None:
                    lista.append(cell.value)
        return lista


for xslx in xslxs:
    for excel_type in config.keys():
            table.append(parseri(xslx, config[excel_type]["sheet"], config[excel_type]["inter_rows"]))

print(table)

table_arr = np.array(table)
Arbeitsplatz = table_arr[::2]

Vorgesetzten =  table_arr[1::2]

print(Arbeitsplatz)

print("and")

print(Vorgesetzten)

"""


wb1 = load_workbook('2019_MUC_22320_Entwicklung_BR&BI_Consulting.xlsx')
# wb2 = load_workbook('2019_MUC_IT22320_Entwicklung_DataScience.xlsx')
sheet = wb1["Mitarbeitergespräch"]
# sheet = wb1["Tabelle1"]
# print (sheet)
table = []
i = 176
for row in sheet.iter_rows(min_row=176, max_row=176, min_col=40, max_col=45):
    #	x = [x.value if x.value is not None  for x in row]
    #	x = [x.value for x in row if x.value is not None]
    #	x = [x.value for x in row]
    i = i + 1
    j = 0
    for cell in row:
        j = j + 1
        if cell.value is not None:
            table.append((cell.value, i, j))

for row in sheet.iter_rows(min_row=186, max_row=186, min_col=40, max_col=45):
    #	x = [x.value if x.value is not None  for x in row]
    #	x = [x.value for x in row if x.value is not None]
    #	x = [x.value for x in row]
    i = i + 1
    j = 0
    for cell in row:
        j = j + 1
        if cell.value is not None:
            table.append((cell.value, i, j))
#	x = [x.value for x in list(row)]
#	x = [v.strip() if isinstance(v, str) else v for v in x]
print(table)
#	x = [v.strip() if isinstance(v, str) or isinstance(v, unicode) else v for v in x]
#	if year_month is not None:
#		x.append(year_month)
#	table.append(x)
xslxs = [f for f in os.listdir(xml_dir) if f.endswith("xslx")]
"""