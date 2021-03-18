"""
Importing necessary modules
"""

import os
import pandas as pd
from openpyxl import load_workbook

# def get_sheet():
#     path = input()
#     list_files = []
#     for f_name in os.listdir(path):
#         if f_name.endswith('.xlsx'):
#             list_files.append(f_name)
#     return list_files
#     return path

def search_data():
    path = input()
    list_files = []
    for f_name in os.listdir(path):
        if f_name.endswith('.xlsx'):
            list_files.append(f_name)
    print(list_files)
    print('Wanna search more?????...')
    ans_more = input()
    while ans_more.lower() == 'yes':
        path = input()
        more_se = []
        for f_name in os.listdir(path):
            if f_name.endswith('.xlsx'):
                more_se.append(f_name)
        more_se.extend(list_files)
        print('Wanna search more?????...')
        ans_more = input()
        if ans_more.lower() == 'no':
            break
    print(list_files)
    for i in range(len(list_files)):
        joiner = "".join(map(str,list_files[i]))
        print(joiner)
        conc = path+"\\"+joiner
        print(conc)
        workbook = load_workbook(filename=conc)
        res = workbook.sheetnames
        print("Total number of sheets in excel file are " + str(len(res)))
        sheets = {}  # empty dictionary
        for i in range(len(res)):
            sheets["Sheet" + str(i)] = pd.read_excel(conc, engine="openpyxl", sheet_name=i)
            sheets["Sheet" + str(i)].dropna(axis=1, how='all', inplace=True)
            print(sheets["Sheet" + str(i)])



search_data()

