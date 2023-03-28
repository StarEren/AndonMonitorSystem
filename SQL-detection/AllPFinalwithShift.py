import requests
import datetime
import openpyxl
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication

#####################################################  DATABASE CONNECTION  #####################################################

# Set up the API endpoints and parameters for each database
url1 = 'http://10.10.12.101/sql-request.do'
params1 = {
    'response_type': 'application/json',
    'sql_statement': 'select reason, start_time from timeline_stream'
}

url2 = 'http://10.10.12.102/sql-request.do'
params2 = {
    'response_type': 'application/json',
    'sql_statement': 'select reason, start_time from timeline_stream'
}

url3 = 'http://10.10.12.103/sql-request.do'
params3 = {
    'response_type': 'application/json',
    'sql_statement': 'select reason, start_time from timeline_stream'
}

# Make the API requests for each database
response1 = requests.post(url1, data=params1)
response2 = requests.post(url2, data=params2)
response3 = requests.post(url3, data=params3)

# Check for errors in each response
for response in [response1, response2, response3]:
    if response.status_code != 200:
        print(f'Request failed with error code {response.status_code}')
        exit()

# Parse the response data for each database as JSON
data1 = response1.json()
data2 = response2.json()
data3 = response3.json()

# Check for errors in each response
for data in [data1, data2, data3]:
    if data[0]['error']:
        print(f'Request failed with error: {data[0]["error"]}')
        exit()

# Extract the data from each response
rows1 = data1[0]['data']
rows2 = data2[0]['data']
rows3 = data3[0]['data']

#####################################################  ERROR ALGORITHM  #####################################################

# Initialize dictionaries to store error times for each database and shift
error_times1 = {}
error_times2 = {}
error_times3 = {}

# Initialize variables to track the current error code, start time, and shift for each database
current_error_code1 = None
current_start_time1 = None
current_shift1 = None
current_error_code2 = None
current_start_time2 = None
current_shift2 = None
current_error_code3 = None
current_start_time3 = None
current_shift3 = None

# Iterate over the rows for each database and calculate the total time for each error for each day and shift
for row in rows1:
    # Extract the data from the response
    error_code = row[0]
    start_time = float(row[1])
    if error_code:
        # Convert the start time to a datetime object
        days, seconds = divmod(start_time, 86400)
        base_date = datetime.datetime(1900, 1, 1)
        timestamp = base_date + datetime.timedelta(days=days, seconds=seconds)

        # Assign shift based on time of error
        if 7 <= timestamp.hour < 15:
            shift = "Day Shift"
        elif 15 <= timestamp.hour < 23:
            shift = "Night Shift"
        else:
            shift = "Afternoon Shift"
        
        # If the error code or shift has changed or this is the last row, calculate the total time for the previous error code and shift
        if error_code != current_error_code1 or shift != current_shift1 or row == rows1[-1]:
            if current_error_code1:
                # Calculate the total time for the current error code and shift
                if current_start_time1.date() != timestamp.date():
                    total_time = (datetime.datetime.combine(current_start_time1.date(), datetime.time.max) - current_start_time1).total_seconds() / 3600
                else:
                    total_time = (timestamp - current_start_time1).total_seconds() / 3600
                
                if current_error_code1 not in error_times1:
                    error_times1[current_error_code1] = {}
                if current_shift1 not in error_times1[current_error_code1]:
                    error_times1[current_error_code1][current_shift1] = {}
                if current_start_time1.date() not in error_times1[current_error_code1][current_shift1]:
                    error_times1[current_error_code1][current_shift1][current_start_time1.date()] = 0
                error_times1[current_error_code1][current_shift1][current_start_time1.date()] += total_time

            # Update the current error code, start time, and shift
            current_error_code1 = error_code
            current_start_time1 = timestamp
            current_shift1 = shift

for row in rows2:
    # Extract the data from the response
    error_code = row[0]
    start_time = float(row[1])
    if error_code:
        # Convert the start time to a datetime object
        days, seconds = divmod(start_time, 86400)
        base_date = datetime.datetime(1900, 1, 1)
        timestamp = base_date + datetime.timedelta(days=days, seconds=seconds)

        # Assign shift based on time of error
        if 7 <= timestamp.hour < 15:
            shift = "Day Shift"
        elif 15 <= timestamp.hour < 23:
            shift = "Night Shift"
        else:
            shift = "Afternoon Shift"
        
        # If the error code or shift has changed or this is the last row, calculate the total time for the previous error code and shift
        if error_code != current_error_code2 or shift != current_shift2 or row == rows2[-1]:
            if current_error_code2:
                # Calculate the total time for the current error code and shift
                if current_start_time2.date() != timestamp.date():
                    total_time = (datetime.datetime.combine(current_start_time2.date(), datetime.time.max) - current_start_time2).total_seconds() / 3600
                else:
                    total_time = (timestamp - current_start_time2).total_seconds() / 3600
                
                if current_error_code2 not in error_times2:
                    error_times2[current_error_code2] = {}
                if current_shift2 not in error_times2[current_error_code2]:
                    error_times2[current_error_code2][current_shift2] = {}
                if current_start_time2.date() not in error_times2[current_error_code2][current_shift2]:
                    error_times2[current_error_code2][current_shift2][current_start_time2.date()] = 0
                error_times2[current_error_code2][current_shift2][current_start_time2.date()] += total_time

            # Update the current error code, start time, and shift
            current_error_code2 = error_code
            current_start_time2 = timestamp
            current_shift2 = shift

for row in rows3:
    # Extract the data from the response
    error_code = row[0]
    start_time = float(row[1])
    if error_code:
        # Convert the start time to a datetime object
        days, seconds = divmod(start_time, 86400)
        base_date = datetime.datetime(1900, 1, 1)
        timestamp = base_date + datetime.timedelta(days=days, seconds=seconds)

        # Assign shift based on time of error
        if 7 <= timestamp.hour < 15:
            shift = "Day Shift"
        elif 15 <= timestamp.hour < 23:
            shift = "Night Shift"
        else:
            shift = "Afternoon Shift"
        
        # If the error code or shift has changed or this is the last row, calculate the total time for the previous error code and shift
        if error_code != current_error_code3 or shift != current_shift3 or row == rows3[-1]:
            if current_error_code3:
                # Calculate the total time for the current error code and shift
                if current_start_time3.date() != timestamp.date():
                    total_time = (datetime.datetime.combine(current_start_time3.date(), datetime.time.max) - current_start_time3).total_seconds() / 3600
                else:
                    total_time = (timestamp - current_start_time3).total_seconds() / 3600
                
                if current_error_code3 not in error_times3:
                    error_times3[current_error_code3] = {}
                if current_shift3 not in error_times3[current_error_code3]:
                    error_times3[current_error_code3][current_shift3] = {}
                if current_start_time3.date() not in error_times3[current_error_code3][current_shift3]:
                    error_times3[current_error_code3][current_shift3][current_start_time3.date()] = 0
                error_times3[current_error_code3][current_shift3][current_start_time3.date()] += total_time

            # Update the current error code, start time, and shift
            current_error_code3 = error_code
            current_start_time3 = timestamp
            current_shift3 = shift

# Print the total time for each error for each day for each database and shift
for db_name, error_times in [("P1", error_times1), ("P2", error_times2), ("P3", error_times3)]:
    print(f"\n{db_name}")
    for error_code, date_times in error_times.items():
        for date, shift_times in sorted(date_times.items(), key=lambda x: x[0]):
            for shift, total_time in shift_times.items():
                print(f'Error code {error_code} on {date} ({shift}): {total_time:.2f} hours')

