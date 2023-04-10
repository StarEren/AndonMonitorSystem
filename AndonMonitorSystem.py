import datetime
import requests
import openpyxl

# print("Please close the excel file in...")
# countdown_time = 10

# while countdown_time > 0:
#     print(countdown_time)
#     time.sleep(1)
#     countdown_time -= 1

# print("Starting Refresh...")

# Define the start and end dates for this month
today = datetime.datetime.now()
start_date = datetime.datetime(today.year, today.month, 1)
end_date = datetime.datetime(today.year, today.month + 1, 1) - datetime.timedelta(days=1)

# Define the start and end times for each shift
day_shift_start = datetime.time(hour=6)
day_shift_end = datetime.time(hour=14)
afternoon_shift_start = datetime.time(hour=14)
afternoon_shift_end = datetime.time(hour=22)
night_shift_start = datetime.time(hour=22)
night_shift_end = datetime.time(hour=6)

# Set up the API endpoint and parameters
url1 = 'http://10.10.12.101/sql-request.do'
url2 = 'http://10.10.12.102/sql-request.do'
url3 = 'http://10.10.12.103/sql-request.do'

params = {
    'response_type': 'application/json',
    'sql_statement': 'select start_time, state, reason, duration from timeline_stream'
}

# Make the API request
response1 = requests.post(url1, data=params)
response2 = requests.post(url2, data=params)
response3 = requests.post(url3, data=params)

# Parse the response data as JSON
data1 = response1.json()
data2 = response2.json()
data3 = response3.json()


# Define the start and end dates for this month
today = datetime.datetime.now()
start_date = datetime.datetime(today.year, today.month, 1)
end_date = datetime.datetime(today.year, today.month + 1, 1) - datetime.timedelta(days=1)

# Define the start and end times for each shift
day_shift_start = datetime.time(hour=6)
day_shift_end = datetime.time(hour=14)
afternoon_shift_start = datetime.time(hour=14)
afternoon_shift_end = datetime.time(hour=22)
night_shift_start = datetime.time(hour=22)
night_shift_end = datetime.time(hour=6)

# Extract the data from the response and filter by date and shift
rows1 = data1[0]['data']
rows2 = data2[0]['data']
rows3 = data3[0]['data']

run_durations1 = {'Day': {}, 'Afternoon': {}, 'Night': {}}
down_durations1 = {'Day': {}, 'Afternoon': {}, 'Night': {}}
run_durations2 = {'Day': {}, 'Afternoon': {}, 'Night': {}}
down_durations2 = {'Day': {}, 'Afternoon': {}, 'Night': {}}
run_durations3 = {'Day': {}, 'Afternoon': {}, 'Night': {}}
down_durations3 = {'Day': {}, 'Afternoon': {}, 'Night': {}}

for row in rows1:
    start_time = int(float(row[0]))
    timestamp = datetime.datetime(1900, 1, 1) + datetime.timedelta(seconds=start_time)
    state = int(row[1])
    state_text = 'Run' if state == 0 else 'Down'
    reason = row[2]
    
    # Check if the start_time falls within a specific shift
    if day_shift_start <= timestamp.time() < day_shift_end:
        shift = 'Day'
    elif afternoon_shift_start <= timestamp.time() < afternoon_shift_end:
        shift = 'Afternoon'
    else:
        shift = 'Night'
        
    if start_date <= timestamp <= end_date:
        duration_in_seconds = int(float(row[3]))
        duration_in_minutes, duration_in_seconds = divmod(duration_in_seconds, 60)
        duration_in_hours, duration_in_minutes = divmod(duration_in_minutes, 60)
        duration_formatted = f'{duration_in_hours:02d}:{duration_in_minutes:02d}:{duration_in_seconds:02d}'
        
        # Check if the state is 0 (Run) or 1 (Down) and add the duration to the corresponding dictionary
        if state == 0:
            run_durations1[shift][reason] = run_durations1[shift].get(reason, {})
            run_durations1[shift][reason][timestamp.date()] = run_durations1[shift][reason].get(timestamp.date(), 0) + float(row[3])
        elif state == 1:
            down_durations1[shift][reason] = down_durations1[shift].get(reason, {})
            down_durations1[shift][reason][timestamp.date()] = down_durations1[shift][reason].get(timestamp.date(), 0) + float(row[3])

for row in rows2:
    start_time = int(float(row[0]))
    timestamp = datetime.datetime(1900, 1, 1) + datetime.timedelta(seconds=start_time)
    state = int(row[1])
    state_text = 'Run' if state == 0 else 'Down'
    reason = row[2]
    
    # Check if the start_time falls within a specific shift
    if day_shift_start <= timestamp.time() < day_shift_end:
        shift = 'Day'
    elif afternoon_shift_start <= timestamp.time() < afternoon_shift_end:
        shift = 'Afternoon'
    else:
        shift = 'Night'
        
    if start_date <= timestamp <= end_date:
        duration_in_seconds = int(float(row[3]))
        duration_in_minutes, duration_in_seconds = divmod(duration_in_seconds, 60)
        duration_in_hours, duration_in_minutes = divmod(duration_in_minutes, 60)
        duration_formatted = f'{duration_in_hours:02d}:{duration_in_minutes:02d}:{duration_in_seconds:02d}'
        # print(f'Start Time: {timestamp}, Shift: {shift}, State: {state_text}, Reason: {reason}, Duration: {duration_formatted}')
        
        # Check if the state is 0 (Run) or 1 (Down) and add the duration to the corresponding dictionary
        if state == 0:
            run_durations2[shift][reason] = run_durations2[shift].get(reason, {})
            run_durations2[shift][reason][timestamp.date()] = run_durations2[shift][reason].get(timestamp.date(), 0) + float(row[3])
        elif state == 1:
            down_durations2[shift][reason] = down_durations2[shift].get(reason, {})
            down_durations2[shift][reason][timestamp.date()] = down_durations2[shift][reason].get(timestamp.date(), 0) + float(row[3])

for row in rows3:
    start_time = int(float(row[0]))
    timestamp = datetime.datetime(1900, 1, 1) + datetime.timedelta(seconds=start_time)
    state = int(row[1])
    state_text = 'Run' if state == 0 else 'Down'
    reason = row[2]
    
    # Check if the start_time falls within a specific shift
    if day_shift_start <= timestamp.time() < day_shift_end:
        shift = 'Day'
    elif afternoon_shift_start <= timestamp.time() < afternoon_shift_end:
        shift = 'Afternoon'
    else:
        shift = 'Night'
        
    if start_date <= timestamp <= end_date:
        duration_in_seconds = int(float(row[3]))
        duration_in_minutes, duration_in_seconds = divmod(duration_in_seconds, 60)
        duration_in_hours, duration_in_minutes = divmod(duration_in_minutes, 60)
        duration_formatted = f'{duration_in_hours:02d}:{duration_in_minutes:02d}:{duration_in_seconds:02d}'
        # print(f'Start Time: {timestamp}, Shift: {shift}, State: {state_text}, Reason: {reason}, Duration: {duration_formatted}')
        
        # Check if the state is 0 (Run) or 1 (Down) and add the duration to the corresponding dictionary
        if state == 0:
            run_durations3[shift][reason] = run_durations3[shift].get(reason, {}) 
            run_durations3[shift][reason][timestamp.date()] = run_durations3[shift][reason].get(timestamp.date(), 0) + float(row[3])
        elif state == 1:
            down_durations3[shift][reason] = down_durations3[shift].get(reason, {})
            down_durations3[shift][reason][timestamp.date()] = down_durations3[shift][reason].get(timestamp.date(), 0) + float(row[3]) 

# print("\n              Date, Reason: Duration (hours)")
# print("\n10.10.12.101\n\n       Run durations:")
# for shift in run_durations1:
#     print(f'         {shift}')
#     for reason, durations in run_durations1[shift].items():
#         for date, duration in durations.items():
#             print(f'              {date}, {reason}: {duration}')
# print('\n       Down durations:  ')
# for shift in down_durations1:
#     print(f'         {shift}')
#     for reason, durations in down_durations1[shift].items():
#         for date, duration in durations.items():
#             print(f'              {date}, {reason}: {duration}')

# print("\n10.10.12.102\n\n       Run durations:")
# for shift in run_durations2:
#     print(f'         {shift}')
#     for reason, durations in run_durations2[shift].items():
#         for date, duration in durations.items():
#             print(f'              {date}, {reason}: {duration}')
# print('\n       Down durations:  ')
# for shift in down_durations2:
#     print(f'         {shift}')
#     for reason, durations in down_durations2[shift].items():
#         for date, duration in durations.items():
#             print(f'              {date}, {reason}: {duration}')

# print("\n10.10.12.103\n\n       Run durations:")
# for shift in run_durations3:
#     print(f'         {shift}')
#     for reason, durations in run_durations3[shift].items():
#         for date, duration in durations.items():
#             print(f'              {date}, {reason}: {duration}')
# print('\n       Down durations:  ')
# for shift in down_durations3:
#     print(f'         {shift}')
#     for reason, durations in down_durations3[shift].items():
#         for date, duration in durations.items():
#             print(f'              {date}, {reason}: {duration}')

# print("\n              Date, Reason: Duration (hours)")
# print("\n10.10.12.101\n\n       Run durations:")
# for shift in run_durations1:
#     print(f'         {shift}')
#     for reason, durations in run_durations1[shift].items():
#         for date, duration in durations.items():
#             hours = duration // 3600
#             minutes = (duration % 3600) // 60
#             seconds = duration % 60
#             print(f'              {date}, {reason}: {int(hours):02d}:{int(minutes):02d}:{int(seconds):02d}')
# print('\n       Down durations:  ')
# for shift in down_durations1:
#     print(f'         {shift}')
#     for reason, durations in down_durations1[shift].items():
#         for date, duration in durations.items():
#             hours = duration // 3600
#             minutes = (duration % 3600) // 60
#             seconds = duration % 60
#             print(f'              {date}, {reason}: {int(hours):02d}:{int(minutes):02d}:{int(seconds):02d}')

# print("\n10.10.12.102\n\n       Run durations:")
# for shift in run_durations2:
#     print(f'         {shift}')
#     for reason, durations in run_durations2[shift].items():
#         for date, duration in durations.items():
#             hours = duration // 3600
#             minutes = (duration % 3600) // 60
#             seconds = duration % 60
#             print(f'              {date}, {reason}: {int(hours):02d}:{int(minutes):02d}:{int(seconds):02d}')
# print('\n       Down durations:  ')
# for shift in down_durations2:
#     print(f'         {shift}')
#     for reason, durations in down_durations2[shift].items():
#         for date, duration in durations.items():
#             hours = duration // 3600
#             minutes = (duration % 3600) // 60
#             seconds = duration % 60
#             print(f'              {date}, {reason}: {int(hours):02d}:{int(minutes):02d}:{int(seconds):02d}')

# print("\n10.10.12.103\n\n       Run durations:")
# for shift in run_durations3:
#     print(f'         {shift}')
#     for reason, durations in run_durations3[shift].items():
#         for date, duration in durations.items():
#             hours = duration // 3600
#             minutes = (duration % 3600) // 60
#             seconds = duration % 60
#             print(f'              {date}, {reason}: {int(hours):02d}:{int(minutes):02d}:{int(seconds):02d}')
# print('\n       Down durations:  ')
# for shift in down_durations3:
#     print(f'         {shift}')
#     for reason, durations in down_durations3[shift].items():
#         for date, duration in durations.items():
#             hours = duration // 3600
#             minutes = (duration % 3600) // 60
#             seconds = duration % 60
#             print(f'              {date}, {reason}: {int(hours):02d}:{int(minutes):02d}:{int(seconds):02d}')


print('---------------')

# Define a dictionary to map day numbers to column letters
day_to_column = {1: 'C', 2: 'D', 3: 'E', 4: 'F', 5: 'G', 6: 'H', 7: 'I', 8: 'J', 9: 'K', 10: 'L',
                 11: 'M', 12: 'N', 13: 'O', 14: 'P', 15: 'Q', 16: 'R', 17: 'S', 18: 'T', 19: 'U',
                 20: 'V', 21: 'W', 22: 'X', 23: 'Y', 24: 'Z', 25: 'AA', 26: 'AB', 27: 'AC', 28: 'AD',
                 29: 'AE', 30: 'AF', 31: 'AG'}

# Define a list of dictionaries, where each dictionary corresponds to a different database
downreason_to_row1 = {'0' : 57, '1000' : 49, '1001' : 52, '1002' : 51, '1003' : 53, '1004' : 50, '1005' : 54, '1006' : 46, '1007' : 45,
                 '1008' : 43, '1009' : 44, '1010' : 47, '1011' : 39, '1012' : 41, '1013' : 37, '1014' : 40, '1015' : 38, '1016' : 55}

downreason_to_row2 = {'0' : 79, '1000' : 71, '1001' : 74, '1002' : 73, '1003' : 75, '1004' : 72, '1005' : 76, '1006' : 68, '1007' : 67,
                 '1008' : 65, '1009' : 66, '1010' : 69, '1011' : 61, '1012' : 63, '1013' : 59, '1014' : 62, '1015' : 60, '1016' : 77}

downreason_to_row3 = {'0' : 101, '1000' : 93, '1001' : 96, '1002' : 95, '1003' : 97, '1004' : 94, '1005' : 98, '1006' : 90, '1007' : 91,
                 '1008' : 87, '1009' : 88, '1010' : 91, '1011' : 83, '1012' : 85, '1013' : 81, '1014' : 84, '1015' : 82, '1016' : 99}

runtotal_to_row1 = {'0' : 6}

runtotal_to_row2 = {'0' : 9}

runtotal_to_row3 = {'0' : 12}

# Define a dictionary to map shift names to sheet names
shift_to_sheet = {"Day" : "Day", "Afternoon" : "Afternoon", "Night" : "Night"}

# Define the path to the Excel file
file_path = 'AndonReport.xlsx'

#Open the workbook
workbook = openpyxl.load_workbook(file_path)

# Iterate over the error times dictionary and update the corresponding cells in the Excel file
for shift in down_durations1:
    for reason, durations in down_durations1[shift].items():
        for date, duration in durations.items():
            sheet_name = shift_to_sheet.get(shift)
            if sheet_name:
                
                worksheet = workbook[sheet_name]

                # Extract the day number from the date
                day_number = date.day

                # Get the column letter for the day number
                day_column = day_to_column[day_number]

                # Get the row number for the error code
                reason_rowfinal1 = downreason_to_row1[reason]

                # Get the cell to write to
                cell = worksheet[f'{day_column}{reason_rowfinal1}']

                # Write the total time to the cell
                cell.value = f"{int(duration//3600):02d}:{int((duration%3600)//60):02d}:{int(duration%60):02d}"

# Iterate over the error times dictionary and update the corresponding cells in the Excel file
for shift in down_durations2:
    for reason, durations in down_durations2[shift].items():
        for date, duration in durations.items():
            sheet_name = shift_to_sheet.get(shift)
            if sheet_name:
                worksheet = workbook[sheet_name]

                # Extract the day number from the date
                day_number = date.day

                # Get the column letter for the day number
                day_column = day_to_column[day_number]

                # Get the row number for the error code
                reason_rowfinal1 = downreason_to_row2[reason]

                # Get the cell to write to
                cell = worksheet[f'{day_column}{reason_rowfinal1}']

                # Write the total time to the cell
                cell.value = f"{int(duration//3600):02d}:{int((duration%3600)//60):02d}:{int(duration%60):02d}"

# Iterate over the error times dictionary and update the corresponding cells in the Excel file
for shift in down_durations1:
    for reason, durations in down_durations3[shift].items():
        for date, duration in durations.items():
            sheet_name = shift_to_sheet.get(shift)
            if sheet_name:
                worksheet = workbook[sheet_name]

                # Extract the day number from the date
                day_number = date.day

                # Get the column letter for the day number
                day_column = day_to_column[day_number]

                # Get the row number for the error code
                reason_rowfinal1 = downreason_to_row3[reason]

                # Get the cell to write to
                cell = worksheet[f'{day_column}{reason_rowfinal1}']

                # Write the total time to the cell
                cell.value = f"{int(duration//3600):02d}:{int((duration%3600)//60):02d}:{int(duration%60):02d}"

                # Save the workbook

# Iterate over the error times dictionary and update the corresponding cells in the Excel file
for shift in run_durations1:
    for reason, durations in run_durations1[shift].items():
        for date, duration in durations.items():
            sheet_name = shift_to_sheet.get(shift)
            if sheet_name:
                
                worksheet = workbook[sheet_name]

                # Extract the day number from the date
                day_number = date.day

                # Get the column letter for the day number
                day_column = day_to_column[day_number]

                # Get the row number for the error code
                reason_rowfinal1 = runtotal_to_row1[reason]

                # Get the cell to write to
                cell = worksheet[f'{day_column}{reason_rowfinal1}']

                # Write the total time to the cell
                cell.value = f"{int(duration//3600):02d}:{int((duration%3600)//60):02d}:{int(duration%60):02d}"

# Iterate over the error times dictionary and update the corresponding cells in the Excel file
for shift in run_durations2:
    for reason, durations in run_durations2[shift].items():
        for date, duration in durations.items():
            sheet_name = shift_to_sheet.get(shift)
            if sheet_name:
                
                worksheet = workbook[sheet_name]

                # Extract the day number from the date
                day_number = date.day

                # Get the column letter for the day number
                day_column = day_to_column[day_number]

                # Get the row number for the error code
                reason_rowfinal1 = runtotal_to_row2[reason]

                # Get the cell to write to
                cell = worksheet[f'{day_column}{reason_rowfinal1}']

                # Write the total time to the cell
                cell.value = f"{int(duration//3600):02d}:{int((duration%3600)//60):02d}:{int(duration%60):02d}"

# Iterate over the error times dictionary and update the corresponding cells in the Excel file
for shift in run_durations3:
    for reason, durations in run_durations3[shift].items():
        for date, duration in durations.items():
            sheet_name = shift_to_sheet.get(shift)
            if sheet_name:
                
                worksheet = workbook[sheet_name]

                # Extract the day number from the date
                day_number = date.day

                # Get the column letter for the day number
                day_column = day_to_column[day_number]

                # Get the row number for the error code
                reason_rowfinal1 = runtotal_to_row3[reason]

                # Get the cell to write to
                cell = worksheet[f'{day_column}{reason_rowfinal1}']

                # Write the total time to the cell
                cell.value = f"{int(duration//3600):02d}:{int((duration%3600)//60):02d}:{int(duration%60):02d}"

# Save the workbook
workbook.save(file_path)

# print("Refresh Completed.")