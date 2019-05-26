''' -------------------------------- Last Modified on Mon, 24 Mai 2019  --------------------------------'''
''' -------------------------------- Author: Markos Gavalas             --------------------------------'''

'''The following scipt is parsing particular informations from .xlsx files and put them into .csv files'''

from openpyxl import load_workbook

'''
config = {
    "travel" : {
        "keyword"  : "reisedaten",
        "table"    : db_master + ".e_travel",
        "schema"   : ["Travel_ID", "Traveler_ID", "Dep_Location", "Dep_Country",
                      "Dep_Time", "Arr_Location", "Arr_Country",
                      "Arr_Time","Risk_Level","yearmonth"],
        "sheet"    : "Sheet1",
        "hdfs"     : False,
        "partition": True
    }
}
'''
############################################################### global variables

# Please insert the directory where the .xlsx files are in
# if nothing is given the program will search for the files that are in the directory of the script
xslx_dir = "C:/Users/gavam/Desktop/use_case"

# Please insert the directory where the .csv files will be created
# if nothing is given the program will search for the files that are in the directory of the script
csv_dir = "C:/Users/gavam/Desktop/use_case"

# Please insert here the Column name output you would like to have
output_cols = ["Zahlenwert von 1-10", "Anmerkungen"]

# Please insert here the excel sheet rows that correspond to the information you want to parse:

wb1 = load_workbook('2019_MUC_22320_Entwicklung_BR&BI_Consulting.xlsx')
# wb2 = load_workbook('2019_MUC_IT22320_Entwicklung_DataScience.xlsx')
sheet = wb1["Mitarbeitergespräch"]
#sheet = wb1["Tabelle1"]
# print (sheet)
table = []
i = 169
for row in sheet.iter_rows(min_row=170, max_row=180):
#	x = [x.value if x.value is not None  for x in row]
#	x = [x.value for x in row if x.value is not None]
#	x = [x.value for x in row]
    i = i + 1
    j = 0
    for cell in row:
        j = j + 1
        if cell.value is not None:
            table.append((cell.value,i,j))

#	x = [x.value for x in list(row)]
#	x = [v.strip() if isinstance(v, str) else v for v in x]
print(table)
#	x = [v.strip() if isinstance(v, str) or isinstance(v, unicode) else v for v in x]
#	if year_month is not None:
#		x.append(year_month)
#	table.append(x)
