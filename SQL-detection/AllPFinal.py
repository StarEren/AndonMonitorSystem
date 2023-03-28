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

# Initialize dictionaries to store error times for each database
error_times1 = {}
error_times2 = {}
error_times3 = {}

# Initialize variables to track the current error code and start time for each database
current_error_code1 = None
current_start_time1 = None
current_error_code2 = None
current_start_time2 = None
current_error_code3 = None
current_start_time3 = None

#####################################################  ERROR ALGORITHM  #####################################################

# Iterate over the rows for each database and calculate the total time for each error for each day
for row in rows1:
    # Extract the data from the response
    error_code = row[0]
    start_time = float(row[1])
    if error_code:
        # Convert the start time to a datetime object
        days, seconds = divmod(start_time, 86400)
        base_date = datetime.datetime(1900, 1, 1)
        timestamp = base_date + datetime.timedelta(days=days, seconds=seconds)

        # If the error code has changed or this is the last row, calculate the total time for the previous error code
        if error_code != current_error_code1 or row == rows1[-1]:
            if current_error_code1:
                # Calculate the total time for the current error code
                if current_start_time1.date() != timestamp.date():
                    total_time = (datetime.datetime.combine(current_start_time1.date(), datetime.time.max) - current_start_time1).total_seconds() / 3600
                else:
                    total_time = (timestamp - current_start_time1).total_seconds() / 3600

                if current_error_code1 not in error_times1:
                    error_times1[current_error_code1] = {}
                if current_start_time1.date() not in error_times1[current_error_code1]:
                    error_times1[current_error_code1][current_start_time1.date()] = 0
                error_times1[current_error_code1][current_start_time1.date()] += total_time

        # Update the current error code and start time
        current_error_code1 = error_code
        current_start_time1 = timestamp

for row in rows2:
    # Extract the data from the response
    error_code = row[0]
    start_time = float(row[1])
    if error_code:
        # Convert the start time to a datetime object
        days, seconds = divmod(start_time, 86400)
        base_date = datetime.datetime(1900, 1, 1)
        timestamp = base_date + datetime.timedelta(days=days, seconds=seconds)

        # If the error code has changed or this is the last row, calculate the total time for the previous error code
        if error_code != current_error_code2 or row == rows2[-1]:
            if current_error_code2:
                # Calculate the total time for the current error code
                if current_start_time2.date() != timestamp.date():
                    total_time = (datetime.datetime.combine(current_start_time2.date(), datetime.time.max) - current_start_time2).total_seconds() / 3600
                else:
                    total_time = (timestamp - current_start_time2).total_seconds() / 3600
                
                if current_error_code2 not in error_times2:
                    error_times2[current_error_code2] = {}
                if current_start_time2.date() not in error_times2[current_error_code2]:
                    error_times2[current_error_code2][current_start_time2.date()] = 0
                error_times2[current_error_code2][current_start_time2.date()] += total_time

            # Update the current error code and start time
            current_error_code2 = error_code
            current_start_time2 = timestamp

for row in rows3:
    # Extract the data from the response
    error_code = row[0]
    start_time = float(row[1])
    if error_code:
        # Convert the start time to a datetime object
        days, seconds = divmod(start_time, 86400)
        base_date = datetime.datetime(1900, 1, 1)
        timestamp = base_date + datetime.timedelta(days=days, seconds=seconds)
        
        # If the error code has changed or this is the last row, calculate the total time for the previous error code
        if error_code != current_error_code3 or row == rows3[-1]:
            if current_error_code3:
                # Calculate the total time for the current error code
                if current_start_time3.date() != timestamp.date():
                    total_time = (datetime.datetime.combine(current_start_time3.date(), datetime.time.max) - current_start_time3).total_seconds() / 3600
                else:
                    total_time = (timestamp - current_start_time3).total_seconds() / 3600
                
                if current_error_code3 not in error_times3:
                    error_times3[current_error_code3] = {}
                if current_start_time3.date() not in error_times3[current_error_code3]:
                    error_times3[current_error_code3][current_start_time3.date()] = 0
                error_times3[current_error_code3][current_start_time3.date()] += total_time

            # Update the current error code and start time
            current_error_code3 = error_code
            current_start_time3 = timestamp


for db_name, error_times in [("P1", error_times1), ("P2", error_times2), ("P3", error_times3)]:
    print(f"\n{db_name}")
    for error_code, date_times in error_times.items():
        for date, total_time in date_times.items():
            print(f'Error code {error_code} on {date}: {total_time:.2f} hours')


# #####################################################  IMPORT TO EXCEL  #####################################################

# # Open the Excel file
# workbook = openpyxl.load_workbook('SQL-detection\AndonReport.xlsx')
# worksheet = workbook.active

# # Define a dictionary to map day numbers to column letters
# day_to_column = {1: 'C', 2: 'D', 3: 'E', 4: 'F', 5: 'G', 6: 'H', 7: 'I', 8: 'J', 9: 'K', 10: 'L',
#                  11: 'M', 12: 'N', 13: 'O', 14: 'P', 15: 'Q', 16: 'R', 17: 'S', 18: 'T', 19: 'U',
#                  20: 'V', 21: 'W', 22: 'X', 23: 'Y', 24: 'Z', 25: 'AA', 26: 'AB', 27: 'AC', 28: 'AD',
#                  29: 'AE', 30: 'AF', 31: 'AG'}

# # Define a list of dictionaries, where each dictionary corresponds to a different database
# error_to_row1 = {'0' : 57, '1000' : 49, '1001' : 52, '1002' : 51, '1003' : 53, '1004' : 50, '1005' : 54, '1006' : 46, '1007' : 45,
#                  '1008' : 43, '1009' : 44, '1010' : 47, '1011' : 39, '1012' : 41, '1013' : 37, '1014' : 40, '1015' : 38, '1016' : 55}

# error_to_row2 = {'0' : 79, '1000' : 71, '1001' : 74, '1002' : 73, '1003' : 75, '1004' : 72, '1005' : 76, '1006' : 68, '1007' : 67,
#                  '1008' : 65, '1009' : 66, '1010' : 69, '1011' : 61, '1012' : 63, '1013' : 59, '1014' : 62, '1015' : 60, '1016' : 77}

# error_to_row3 = {'0' : 99, '1000' : 91, '1001' : 94, '1002' : 93, '1003' : 95, '1004' : 92, '1005' : 96, '1006' : 88, '1007' : 87,
#                  '1008' : 85, '1009' : 86, '1010' : 89, '1011' : 81, '1012' : 83, '1013' : 79, '1014' : 82, '1015' : 80, '1016' : 97}

# # Iterate over the error times dictionary and update the corresponding cells in the Excel file
# for error_code, date_times in error_times1.items():
#     for date, total_time in date_times.items():
#         # Extract the day number from the date
#         day_number = date.day

#         # Get the column letter for the day number
#         day_column = day_to_column[day_number]

#         # Get the row number for the error code
#         error_rowfinal1 = error_to_row1[error_code]

#         # Get the cell to write to
#         cell = worksheet[f'{day_column}{error_rowfinal1}']

#         # Write the total time to the cell
#         cell.value = total_time

# for error_code, date_times in error_times2.items():
#     for date, total_time in date_times.items():
#         # Extract the day number from the date
#         day_number = date.day

#         # Get the column letter for the day number
#         day_column = day_to_column[day_number]

#         # Get the row number for the error code
#         error_rowfinal2 = error_to_row2[error_code]

#         # Get the cell to write to
#         cell = worksheet[f'{day_column}{error_rowfinal2}']

#         # Write the total time to the cell
#         cell.value = total_time

# for error_code, date_times in error_times3.items():
#     for date, total_time in date_times.items():
#         # Extract the day number from the date
#         day_number = date.day

#         # Get the column letter for the day number
#         day_column = day_to_column[day_number]

#         # Get the row number for the error code
#         error_rowfinal3 = error_to_row3[error_code]

#         # Get the cell to write to
#         cell = worksheet[f'{day_column}{error_rowfinal3}']

#         # Write the total time to the cell
#         cell.value = total_time

# # Save the Excel file
# workbook.save('SQL-detection\AndonReport.xlsx')

# #####################################################  MAKE EMAIL  #####################################################

# # # create message object instance
# # msg = MIMEMultipart()
# # to_list = ["brian.rankin@martinrea.com"]
# # # setup the parameters of the message
# # password = "wmjurkjfflnpltjk"   # VERY SECURE
# # msg['From'] = "andonreportshydroform@gmail.com"
# # msg['To'] = ",".join(to_list)
# # msg['Subject'] = "Andon Report"

# # # add in the message body
# # msg.attach(MIMEText("Here is the andon report:"))

# # # attach a file
# # with open("SQL-detection\AndonReport.xlsx", "rb") as f:
# #     attach = MIMEApplication(f.read(),_subtype="txt")
# #     attach.add_header('Content-Disposition','attachment',filename=str("AndonReport.xlsx"))
# #     msg.attach(attach)

# # # create server
# # server = smtplib.SMTP('smtp.gmail.com', 587)

# # # start TLS for security
# # server.starttls()

# # # Login
# # server.login(msg['From'], password)

# # # send the message via the server.
# # server.sendmail(msg['From'], to_list, msg.as_string())

# # # terminate the SMTP session
# # server.quit()

# # print("ALL TASKS COMPLETED. EMAIL SENT.")