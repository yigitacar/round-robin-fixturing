import pandas as pd

# Read the file as an object
xl = pd.ExcelFile('entries.xlsx')
# Print sheet names as list
print(xl.sheet_names)

# Read the entries
df = pd.read_excel('entries.xlsx')


