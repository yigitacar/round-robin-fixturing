import pandas as pd
from participants import Participants
from groups import Groups
from openpyxl import load_workbook

file_path = 'entries.xlsx'

# Read the file as an object and read the entries
xl = pd.ExcelFile(file_path)
df = pd.read_excel(xl)

# Load an existing workbook
wb = load_workbook(file_path)
# Create a new sheet
ws1 = wb.create_sheet(title='Group Stage')

num_sheets = len(xl.sheet_names)
if num_sheets == 1:
    num_groups = int(input("How many groups?"))
else:
    num_groups = 4

prt = Participants(df)
grp = Groups(num_groups, prt.list_participants)

ws1 = grp.print_groups(ws1)
wb.save('entries.xlsx')



