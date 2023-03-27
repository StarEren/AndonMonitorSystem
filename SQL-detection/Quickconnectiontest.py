import requests
import datetime

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

# Initialize dictionary to store error timestamps
error_timestamps = {}

for row in rows:
    # Extract the data from the response
    error_code = row[0]
    start_time = float(row[1])
    if error_code:
        days, seconds = divmod(start_time, 86400)
        base_date = datetime.datetime(1900, 1, 1)
        timestamp = base_date + datetime.timedelta(days=days, seconds=seconds)
        print(f'{error_code}: {timestamp}')