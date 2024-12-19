import csv

# Set the input and output file paths
input_file = 'PastWinners.csv'  # Replace with the path to your input CSV file
output_file = 'output_sorted.csv'  # Replace with the desired path for the sorted output CSV file

# Read the CSV file and store its data, converting names and emails to lowercase
data = []
with open(input_file, 'r', newline='') as csvfile:
    reader = csv.DictReader(csvfile)
    for row in reader:
        data.append({'Name': row['Name'].strip().lower(), 'Email': row['Email'].strip().lower()})

# Sort the data alphabetically based on the lowercase 'Name' column
sorted_data = sorted(data, key=lambda x: x['Name'])

# Write the sorted data to the output CSV file
with open(output_file, 'w', newline='') as csvfile:
    fieldnames = ['Name', 'Email']  # Replace with the actual field names in your CSV
    writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
    
    # Write the header row
    writer.writeheader()
    
    # Write the sorted rows
    for row in sorted_data:
        writer.writerow({'Name': row['Name'], 'Email': row['Email']})  # Replace with the actual field names in your CSV

print(f"CSV file sorted by 'Name' (case-insensitive) has been saved to {output_file}")
