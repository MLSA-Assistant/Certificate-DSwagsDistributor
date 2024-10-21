import pandas as pd

# Load the first CSV file
file1 = pd.read_csv('LuckyWinners.csv')

# Load the second CSV file
file2 = pd.read_csv('LinkPD.csv')

# Assuming the structure of both CSV files is the same
# Append the second column of file2 to the end of the first column of file1
file1['Link'] = file2['Link']

# Write the merged data to an Excel file
file1.to_excel('Merged.xlsx', index=False)
