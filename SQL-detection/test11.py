import requests
import datetime
import openpyxl

# Set up the API endpoint and parameters
url = 'http://10.10.12.103/sql-request.do'

params = {
    'response_type': 'application/json',
    'sql_statement': 'select reason, start_time from timeline_stream'
}

# Make the API request
response = requests.post(url, data=params)

# Check for errors in the response
if response.status_code != 200:
    print(f'Request failed with error code {response.status_code}')
    exit()

# Parse the response data as JSON
data = response.json()

# Check for errors in the response
if data[0]['error']:
    print(f'Request failed with error: {data[0]["error"]}')
    exit()

# Extract the data from the response
rows = data[0]['data']

# Initialize dictionary to store error times
error_times = {}

# Initialize variables to track the current error code and start time
current_error_code = None
current_start_time = None

# Iterate over the rows and calculate the total time for each error for each day
for row in rows:
    # Extract the data from the response
    error_code = row[0]
    start_time = float(row[1])
    if error_code:
        # Convert the start time to a datetime object
        days, seconds = divmod(start_time, 86400)
        base_date = datetime.datetime(1900, 1, 1)
        timestamp = base_date + datetime.timedelta(days=days, seconds=seconds)

        # If the error code has changed or this is the last row, calculate the total time for the previous error code
        if error_code != current_error_code or row == rows[-1]:
            if current_error_code:
                # Calculate the total time for the current error code
                if current_start_time.date() != timestamp.date():
                    total_time = (datetime.datetime.combine(current_start_time.date(), datetime.time.max) - current_start_time).total_seconds() / 3600
                else:
                    total_time = (timestamp - current_start_time).total_seconds() / 3600
                
                if current_error_code not in error_times:
                    error_times[current_error_code] = {}
                if current_start_time.date() not in error_times[current_error_code]:
                    error_times[current_error_code][current_start_time.date()] = 0
                error_times[current_error_code][current_start_time.date()] += total_time

            # Update the current error code and start time
            current_error_code = error_code
            current_start_time = timestamp

# Print the total time for each error for each day
for error_code, date_times in error_times.items():
    for date, total_time in date_times.items():
        print(f'Error code {error_code} on {date}: {total_time:.2f} hours')

# Open the Excel file
workbook = openpyxl.load_workbook('SQL-detection\Copy of Brian Downtime Report 2018 (002) Eren.xlsx')
worksheet = workbook.active

# Define a dictionary to map day numbers to column letters
day_to_column = {1: 'C', 2: 'D', 3: 'E', 4: 'F', 5: 'G', 6: 'H', 7: 'I', 8: 'J', 9: 'K', 10: 'L',
                 11: 'M', 12: 'N', 13: 'O', 14: 'P', 15: 'Q', 16: 'R', 17: 'S', 18: 'T', 19: 'U',
                 20: 'V', 21: 'W', 22: 'X', 23: 'Y', 24: 'Z', 25: 'AA', 26: 'AB', 27: 'AC', 28: 'AD',
                 29: 'AE', 30: 'AF', 31: 'AG'}

# Define a dictionary to map error codes to row numbers
error_to_row = {0 : 99, 1000: 91, 1001: 94, 1002: 93, 1003 : 95, 1004: 92, 1005 : 96, 1006 : 88, 1007 : 87,
                 1008 : 85, 1009 : 86, 1010 : 89, 1011 : 81, 1012 : 83, 1013 : 79, 1014 : 82, 1015 : 80, 1016 : 97}

# Iterate over the error times dictionary and update the corresponding cells in the Excel file
for error_code, date_times in error_times.items():
    for date, total_time in date_times.items():
        # Extract the day number from the date
        day_number = date.day

        # Get the column letter for the day number
        day_column = day_to_column[day_number]

        # Get the row number for the error code
        error_row = error_to_row[error_code]

        # Get the cell to write to
        cell = worksheet[f'{day_column}{error_row}']

        # Write the total time to the cell
        cell.value = total_time

# Save the Excel file
workbook.save('SQL-detection\Copy of Brian Downtime Report 2018 (002) Eren.xlsx')