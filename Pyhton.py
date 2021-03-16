import pandas as pd
import numpy as np


sheet_1 = pd.read_excel('Book1.xlsx', sheet_name = 0)
sheet_2 = pd.read_excel('Book1.xlsx', sheet_name = 1)
sheet_3 = pd.read_excel('Book1.xlsx', sheet_name = 2)
sheet_4 = pd.read_excel('Book1.xlsx', sheet_name = 3)

all_data_1 = pd.merge(sheet_1, sheet_2, how='left')
all_data_2 = pd.merge(all_data_1, sheet_3, how='left')
all_data_3 = pd.merge(all_data_2, sheet_4, how='left')

# df = pd.read_excel('Book1.xlsx', sheet_name = 'mastersheet')
name = input("Enter the name : ")

listOfColumns = ["Display Name","PS number","Colour","Height","School","Sex","Age","Famsize","Mjob","Fjob",
                 "Car","Room1","Reason","Guardian","Traveltime","Studytime","Failures","Schoolsup",
                 "Group","Date","Activities","Nursery","Higher","Internet","Romantic","Absences",
                 "Class","Weight","SchoolName","Address","City","Postal","Phone"]

df1 = pd.DataFrame(all_data_3, columns = listOfColumns) 
df1.set_index("Display Name",inplace = True)

result = df1.loc[name]


from openpyxl import load_workbook

path = r"Book1.xlsx"
book = load_workbook(path)
writer = pd.ExcelWriter(path, engine = 'openpyxl')
writer.book = book
if 'mastersheet' in book.sheetnames:
    pfd = book['mastersheet'] 
    book.remove(pfd)
    result.to_excel(writer, sheet_name = 'mastersheet')   

writer.save()
writer.close()

result

