""" --------------------------------------  Last Modified on Fri, 07 Jun 2019  --------------------------------------"""

'''---------------------------------------  Author: Markos Gavalas             --------------------------------------'''


'''The following scipt is parsing particular informations from .xlsx files and put them into .csv files              '''
'''Please, close the Excel files that you want to parse before you run the program (otherwise it will give an error) '''

from openpyxl import load_workbook
import sys
import os
import numpy as np
from datetime import date
try:
    from StringIO import StringIO
except ImportError:
    from io import StringIO
#################################################################### configs

'''
 Please insert here:
 1. The excel Table name
 2. The schema of the Table
 3. The sheets
 4. The excel sheet rows 
    (if one cell corresponds to many indices(rows) insert the min index) 
'''

config = {
    "Zufriedenheit mit dem Arbeitsplatz": {
        "schema": ["Zahlenwert von 1-10", "Anmerkungen"],
        "sheet": "Mitarbeitergespräch",
        "inter_rows": [176, 40, 45]
    },
       "Zufriedenheit mit der Betreuung durch den Vorgesetzten": {
        "schema": ["Zahlenwert von 1-10", "Anmerkungen"],
        "sheet": "Mitarbeitergespräch",
        "inter_rows": [186, 40, 45]
    }
}

stamconf = {
    "Stammdaten Stelle": {
        "schema": "KSTFachbereichAbteilung",
        "sheet": "Stammdaten",
        "inter_rows": [5, 4, 9]
    }
}
''' Please insert the directory where the .xlsx files are in. If nothing is given
 the program will search for the files that are in the directory of the script'''
xlsx_dir = "."

'''This will be the directory where the program will put the parsed xlsx files
it will be in parsed_files and under the date where the execution of the program happend'''
today = date.today()

date_folder = today.strftime("%d_%m_%Y")
move_path = "./parsed_files/" + date_folder
try:
    os.makedirs(move_path)
except OSError:
    print("Creation of the directory %s failed" % move_path)
else:
    print("Successfully created the directory %s " % move_path)


'''The program will create 2 .csv files in the directory of the script'''
'''The delimiter will be ;'''
delimiter = ";"
############################################################################
xlsxs = [f for f in os.listdir(xlsx_dir) if f.endswith("xlsx")]


#Function that parse the information into a list
def parser(wb, SheetName, minmaxr, mincol, maxcol):
    sheet = wb[SheetName]
    list_row = []
    for row in sheet.iter_rows(min_row=minmaxr, max_row=minmaxr, min_col=mincol, max_col=maxcol):
        for cell in row:
            if cell.value is not None:
                list_row.append(cell.value.replace(";", "") if isinstance(cell.value,str) else cell.value)
    return list_row


#create a list of wb so that we do not read the excels more that once
wbs = [load_workbook(xlsx) for xlsx in xlsxs]


for excel_type in config.keys():
    lofl = []#list of lists
    for wb in wbs:
        lofl.append(parser(wb, config[excel_type]["sheet"],
                           config[excel_type]["inter_rows"][0],
                           config[excel_type]["inter_rows"][1],
                           config[excel_type]["inter_rows"][2])
                    +
                    parser(wb, stamconf["Stammdaten Stelle"]["sheet"],
                           stamconf["Stammdaten Stelle"]["inter_rows"][0],
                           stamconf["Stammdaten Stelle"]["inter_rows"][1],
                           stamconf["Stammdaten Stelle"]["inter_rows"][2])
                    )
    name = excel_type.replace(" ", "_") + '.csv'
    header = delimiter.join(config[excel_type]["schema"] + [stamconf["Stammdaten Stelle"]["schema"]])
    np.savetxt(name, np.array(lofl), delimiter=delimiter, header=header, comments='', fmt='%s')
    checker = len(lofl)

if checker == len(xlsxs):
    for xslx in xlsxs:
        current_file = "./" + xslx
        destination_file = move_path + "/" + xslx
        os.rename(current_file, destination_file)


