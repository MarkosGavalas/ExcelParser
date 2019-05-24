''' -------------------------------- Last Modified on Mon, 24 Mai 2019  --------------------------------'''
''' -------------------------------- Author: Markos Gavalas             --------------------------------'''

'''The following scipt is parsing particular informations from .xlsx files and put them into .csv files'''


from openpyxl import load_workbook


############################################################### global variables

# Please insert the directory where the .xlsx files are in
# if nothing is given the program will search for the files that are in the directory of the script
xslx_dir = "C:/Users/gavam/Desktop/use_case"

# Please insert the directory where the .csv files will be created
# if nothing is given the program will search for the files that are in the directory of the script
csv_dir = "C:/Users/gavam/Desktop/use_case"

#Please insert here the Column name output you would like to have
output_cols = ["Zahlenwert von 1-10", "Anmerkungen"]

# Please insert here the excel sheet rows that correspond to the information you want to parse:

wb1 = load_workbook('2019_MUC_22320_Entwicklung_BR&BI_Consulting.xlsx')
#wb2 = load_workbook('2019_MUC_IT22320_Entwicklung_DataScience.xlsx')

print (wb1)