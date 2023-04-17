import datetime
import requests
import openpyxl

############################################# DATABASE CONNECTION AND QUERY #############################################

# Set up the API endpoint and parameters
urls = [
    'http://10.10.12.101/sql-request.do',
    #'http://10.10.12.102/sql-request.do',
    'http://10.10.12.103/sql-request.do',
]

params = {
    'response_type': 'application/json',
    'sql_statement': 'select start_time, state, reason, duration from timeline_stream'
}

######################################## DATA TRANSFORMATION AND SORTING ALGORITHM ########################################

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

#Define dictionaries
run_durations = [{'Day': {}, 'Afternoon': {}, 'Night': {}} for _ in range(3)]
down_durations = [{'Day': {}, 'Afternoon': {}, 'Night': {}} for _ in range(3)]

# Define the down_count dictionary
down_count = {}

for i in range(2):

    # Make the API request
    response = requests.post(urls[i], data=params)

    # Parse the response data as JSON
    data = response.json()

    # Extract the data from the response
    rows = data[0]['data']

    for row in rows:

        # Initialize the data
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

            # Check if the state is 0 (Run) or 1 (Down) and add the duration to the corresponding dictionary
            if state == 0:
                run_durations[i][shift][reason] = run_durations[i][shift].get(reason, {})
                run_durations[i][shift][reason][timestamp.date()] = run_durations[i][shift][reason].get(timestamp.date(), 0) + float(row[3])
            elif state == 1:
                down_durations[i][shift][reason] = down_durations[i][shift].get(reason, {})
                down_durations[i][shift][reason][timestamp.date()] = down_durations[i][shift][reason].get(timestamp.date(), 0) + 55
                down_count[shift] = down_count.get(shift, {})
                down_count[shift][reason] = down_count[shift].get(reason, {})
                down_count[shift][reason][timestamp.date()] = down_count[shift][reason].get(timestamp.date(), 0) + 1

####################################################### PRINT DATA #######################################################

# Loop through each IP address
for url in urls:

    # Get the index of the IP address in the urls list
    i = urls.index(url)

    # Print the IP address header and Run durations header
    print(f"\n{url}\n\n       Run durations:")

    # Loop through each shift and reason, and print the run durations
    for shift in run_durations[i]:
        print(f'         {shift}')
        for reason, durations in run_durations[i][shift].items():
            for date, duration in durations.items():
                print(f'              {date}, {reason}: {duration}')

    # Print the down durations header
    print('\n       Down durations:  ')

    # Loop through each shift and reason, and print the down durations and count of occurrences
    for shift in down_durations[i]:
        print(f'         {shift}')
        for reason, durations in down_durations[i][shift].items():
            for date, duration in durations.items():
                print(f'              {date}, {reason}: {duration}')
            for date, count in down_count[shift][reason].items():
                print(f'              {date}, Count of {reason} occurrences: {count}')

################################################## EXPORT TO EXCEL ##################################################

# Define the path to the Excel file
file_path = 'AndonReport.xlsx'

# Define a dictionary to map day numbers to column letters
durationday_to_column = {1: 'C', 2: 'E', 3: 'G', 4: 'I', 5: 'K', 6: 'M', 7: 'O', 8: 'Q', 9: 'S', 10: 'U',
                    11: 'W', 12: 'Y', 13: 'AA', 14: 'AC', 15: 'AE', 16: 'AG', 17: 'AI', 18: 'AK', 19: 'AM',
                    20: 'AO', 21: 'AQ', 22: 'AS', 23: 'AU', 24: 'AW', 25: 'AY', 26: 'BA', 27: 'BC', 28: 'BE',
                    29: 'BG', 30: 'BI', 31: 'BK'}

countday_to_column = {1: 'D', 2: 'F', 3: 'H', 4: 'J', 5: 'L', 6: 'N', 7: 'P', 8: 'R', 9: 'T', 10: 'V',
                    11: 'X', 12: 'Z', 13: 'AB', 14: 'AD', 15: 'AF', 16: 'AH', 17: 'AJ', 18: 'AL', 19: 'AN',
                    20: 'AP', 21: 'AR', 22: 'AT', 23: 'AV', 24: 'AX', 25: 'AZ', 26: 'BB', 27: 'BD', 28: 'BF',
                    29: 'BH', 30: 'BJ', 31: 'BL'}


# Define a list of dictionaries to map reasons, where each dictionary corresponds to a different database
downreason_to_row = [
    {'0': 57, '1000': 49, '1001': 52, '1002': 51, '1003': 53, '1004': 50, '1005': 54, '1006': 46, '1007': 45, '1008': 43, '1009': 44, '1010': 47, '1011': 39, '1012': 41, '1013': 37, '1014': 40, '1015': 38, '1016': 55},
    #{'0': 79, '1000': 71, '1001': 74, '1002': 73, '1003': 75, '1004': 72, '1005': 76, '1006': 68, '1007': 67, '1008': 65, '1009': 66, '1010': 69, '1011': 61, '1012': 63, '1013': 59, '1014': 62, '1015': 60, '1016': 77},
    {'0': 101, '1000': 93, '1001': 96, '1002': 95, '1003': 97, '1004': 94, '1005': 98, '1006': 90, '1007': 91, '1008': 87, '1009': 88, '1010': 91, '1011': 83, '1012': 85, '1013': 81, '1014': 84, '1015': 82, '1016': 99}
]

runtotal_to_row = [
    {'0': 6},
    #{'0': 9},
    {'0': 12}
]

# Define a dictionary to map shift names to sheet names
shift_to_sheet = {"Day": "Day", "Afternoon": "Afternoon", "Night": "Night"}

# Open the workbook
workbook = openpyxl.load_workbook(file_path)

# Iterate over the databases, shifts, and reasons
for i in range(2):
    for shift in ["Day", "Afternoon", "Night"]:
        for reason in downreason_to_row[i]:

            # Get the sheet for the current shift
            worksheet = workbook[shift_to_sheet[shift]]

            # Get the row for the current reason
            reason_row = downreason_to_row[i][reason]

            # Iterate over the dates and durations in down_durations
            for date, duration in down_durations[i][shift].get(reason, {}).items():

                # Get the column for the current date
                day_column = durationday_to_column[date.day]
                
                # Seconds to hours
                hoursduration = duration / 3600
                
                # Combine the column and row
                cell = worksheet[f'{day_column}{reason_row}']

                # Update the cell with the duration
                cell.value = hoursduration

            # Iterate over the dates and durations in run_durations
            for date, duration in run_durations[i][shift].get(reason, {}).items():

                # Get the row for the current reason
                reason_row = runtotal_to_row[i][reason]

                # Get the column for the current date
                day_column = durationday_to_column[date.day]

                # Seconds to hours
                hoursduration = duration / 3600
                
                # Combine the column and row
                cell = worksheet[f'{day_column}{reason_row}']

                # Update the cell with the duration
                cell.value = hoursduration

            # Iterate over the dates and counts in down_count
            for date, count in down_count[shift].get(reason, {}).items():

                # Get the row for the current reason
                reason_row = downreason_to_row[i][reason]

                # Get the column for the current date
                count_column = countday_to_column[date.day]

                # Combine the column and row
                cell = worksheet[f'{count_column}{reason_row}']

                # Update the cell with the count
                cell.value = int(count)
            
            
            
# Save the workbook
workbook.save(file_path)