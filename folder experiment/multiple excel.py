"""
Importing necessary modules
"""

import os
import pandas as pd
from openpyxl import load_workbook

def get_sheet():
    path = input()
    list_files = []
    for f_name in os.listdir(path):
        if f_name.endswith('.xlsx'):
            list_files.append(f_name)
    return list_files

def search_data():
    searched_data = get_sheet()
    print(searched_data)
    print('Wanna search more?????...')
    ans_more = input()
    while ans_more.lower() == 'yes':
        more_se = get_sheet()
        searched_data.extend(more_se)
        print('Wanna search more?????...')
        ans_more = input()
        if ans_more.lower() == 'no':
            break
    return searched_data
search_data()

