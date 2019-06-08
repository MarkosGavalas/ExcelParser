""" --------------------------------------  Last Modified on Fri, 07 Jun 2019  --------------------------------------"""

'''---------------------------------------  Author: Markos Gavalas             --------------------------------------'''


'''The following scipt is parsing particular informations from .xlsx files and put them into .csv files              '''
'''Please, close the Excel files that you want to parse before you run the program (otherwise it will give an error) '''

from openpyxl import load_workbook
import os
import numpy as np
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
        "schema": ["Z.ahlenwert von 1-10", "Anmerkungen"],
        "sheet": "Mitarbeitergespräch",
        "inter_rows": 176
    },
       "Zufriedenheit mit der Betreuung durch den Vorgesetzten": {
        "schema": ["Zahlenwert von 1-10", "Anmerkungen"],
        "sheet": "Mitarbeitergespräch",
        "inter_rows": 186
    }
}
''' Please insert the directory where the .xlsx files are in. If nothing is given
 the program will search for the files that are in the directory of the script'''
xlsx_dir = "."

'''The program will create 2 .csv files in the directory of the script'''
'''The delimiter will be ;'''
delimiter = ";"
############################################################################

xlsxs = [f for f in os.listdir(xlsx_dir) if f.endswith("xlsx")]


#Function that parse the information into a list
def parser(xlsx, SheetName, minmaxr):
    wb = load_workbook(xlsx)
    sheet = wb[SheetName]
    list_row = []
    for row in sheet.iter_rows(min_row=minmaxr, max_row=minmaxr, min_col=40, max_col=45):
        for cell in row:
            if cell.value is not None:
                list_row.append(cell.value.replace(";", "") if isinstance(cell.value,str) else cell.value)
    return list_row


for excel_type in config.keys():
    lofl = []#list of lists
    for xslx in xlsxs:
        lofl.append(parser(xslx, config[excel_type]["sheet"], config[excel_type]["inter_rows"]))
    name = excel_type.replace(" ", "_") + '.csv'
    header = delimiter.join(config[excel_type]["schema"])
    np.savetxt(name, np.array(lofl), delimiter=delimiter, header=header, comments='', fmt='%s')



