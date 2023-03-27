import datetime
import requests

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

# Initialize dictionary to store error totals
error_totals = {}

# Initialize timestamp for first row
last_timestamp = None

# Loop through rows and add up the total time for each error code
for row in rows:
    error_code = row[0]
    timestamp = row[1]
    if error_code:
        if last_timestamp:
            time_diff = (float(timestamp) - float(last_timestamp)) / 3600
            if error_code in error_totals:
                error_totals[error_code] += time_diff
            else:
                error_totals[error_code] = time_diff
        last_timestamp = timestamp

# Print the error totals
if error_totals:
    print("Error totals:")
    for error_code, total_time in error_totals.items():
        print(f"{error_code}: {total_time:.2f} hours")
else:
    print("No error data found.")
