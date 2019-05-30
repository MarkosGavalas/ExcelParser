''' -------------------------------- Last Modified on Mon, 24 Mai 2019  --------------------------------'''
from typing import List, Any

''' -------------------------------- Author: Markos Gavalas             --------------------------------'''

'''The following scipt is parsing particular informations from .xlsx files and put them into .csv files'''
'''It is important to close the Excel files that you want to parse before you run the program (otherwise it will give an error)'''

from openpyxl import load_workbook
import os
import numpy as np

try:
    from StringIO import StringIO
except ImportError:
    from io import StringIO

############################################################## configs

'''Please insert here the excel sheet rows that correspond to the information you want to parse:'''

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
''' Please insert the directory where the .xlsx files are in if 
nothing is given the program will search for the files that are in the directory of the script'''

xslx_dir = "."

'''The program will create 2 .csv files in the directory of the script'''
'''The delimiter will be ;'''
delimiter = ";"
header = delimiter.join( config["Zufriedenheit mit dem Arbeitsplatz"]["schema"])
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
Vorgesetzten = table_arr[1::2]

np.savetxt("Arbeitsplatz.csv", Arbeitsplatz, delimiter=delimiter, header=header, comments='', fmt='%s')
np.savetxt("Vorgesetzten.csv", Vorgesetzten, delimiter=delimiter, header=header, comments='', fmt='%s')

print(Arbeitsplatz)

print("and")

print(Vorgesetzten)
