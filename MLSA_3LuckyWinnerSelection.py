import csv
import random

# Set the input and output file paths
input_file = 'completed_modules.csv'  # Replace with the path to your input CSV file
output_file = 'LuckyWinners.csv'  # Replace with the desired path for the sorted output CSV file
provided_csv_file = 'PastWinners.csv'  # Replace with the path to the provided CSV file

# Read the provided CSV file to get a list of existing email addresses, lowercase names, and lowercase first names
existing_entries = set()
existing_first_names = set()
with open(provided_csv_file, 'r', newline='') as provided_csv:
    reader = csv.DictReader(provided_csv)
    for row in reader:
        existing_entries.add((row['Name'].strip().lower(), row['Email'].strip().lower()))
        existing_first_names.add(row['Name'].strip().split()[0].lower())

# Read the original CSV file and store its data, converting names, emails, and first names to lowercase
data = []
with open(input_file, 'r', newline='') as csvfile:
    reader = csv.DictReader(csvfile)
    for row in reader:
        # Exclude entries with only the first name
        name = row['Name'].strip().lower()
        email = row['Email'].strip().lower()
        first_name = name.split()[0]
        data.append((name, email, first_name))

# Shuffle the data to randomize the order
random.shuffle(data)

# Randomly select 25 unique rows from the shuffled data, checking for duplicates with past winners and first name duplicates
selected_rows = []
while len(selected_rows) < 5 and data:
    candidate = data.pop()  # Pop the last item for random selection
    name, email, first_name = candidate
    if name not in existing_entries and first_name not in existing_first_names:
        selected_rows.append(candidate)
        existing_entries.add(name)
        existing_first_names.add(first_name)

# Sort the selected rows alphabetically by 'Name'
selected_rows = sorted(selected_rows, key=lambda x: x[0])

# Write the selected and sorted rows to the output CSV file, ensuring no duplicates and no matching first names
with open(output_file, 'w', newline='') as csvfile:
    fieldnames = ['Name', 'Email']
    writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
    
    # Write the header row
    writer.writeheader()
    
    # Write the selected rows
    for row in selected_rows:
        writer.writerow({'Name': row[0], 'Email': row[1]})

print("Randomly selected and sorted 25 names and unique emails have been saved to", output_file)
