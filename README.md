# Winter 2023 Co-op WorkTerm Project: Andon Report System
The Andon Report System is a Python script designed to fetch, process, and analyze machine downtime data from multiple databases using API endpoints. The primary goal of the project is to generate a comprehensive report on machine downtime durations and occurrences, sorted by shift and reason, for a specified time range. The script outputs the processed data both in the console and in an Excel file, providing an easy-to-understand and organized overview of machine performance.

## Contents
1. **'AndonReport.py'**: The main Python script that fetches, processes, and exports data to an Excel file.
2. **'AndonReport.xlsx'**: The Excel file where the report data is exported.
3. **'Installation_RunAsAdmin.bat'**: A batch file to check for Python installation and install it if needed, along with required Python modules.
4. **'Refresh.bat'**: A batch file to run the AndonReport.py script and open the generated Excel file.
5. **'XLReasonCodesConfiguration.csv'**: The CSV file contains a list of IDs, states, and user-defined reasons for events in the Andon system.

## Installation
1. Clone or download this repository to your local machine.
2. Run the Rfresh.bat file to execute the AndonReport.py script, which will fetch data, process it, and export it to the AndonReport.xlsx file. The Excel file will open automatically once the process is complete.

## How it works
The AndonReport.py script performs the following tasks:

#### Data Acquisition
  1. The script starts by importing the necessary Python modules, such as datetime, requests, and openpyxl.
  2. It establishes a connection to the specified API endpoints by creating custom parameters with the SQL statement and making a POST request to the appropriate URL using the requests module.
  3. Upon successful connection, it fetches raw data in JSON format and loads it into a list of rows.
#### Data Transformation and Processing
  1. The script processes the raw data by defining the start and end dates for the current month and setting up the start and end times for each shift.
  2. It initializes dictionaries to store run and down durations for each machine and shift, as well as a dictionary to store the count of down occurrences.
  3. The script loops through the data rows, converting the start_time to a timestamp and checking if it falls within a specific shift.
  4. If the timestamp is within the specified date range, it checks if the state is 0 (Run) or 1 (Down) and adds the corresponding duration to the appropriate              dictionary. For down states, it also increments the count of occurrences.
#### Export to Excel
  1. The script defines the path to the Excel file and initializes dictionaries to map day numbers to column letters, reasons to row numbers, and shifts to sheet            names.
  2. It loads the Excel workbook and iterates over the databases, shifts, and reasons.
  3. For each combination of database, shift, and reason, it updates the corresponding cell in the Excel workbook with the run and down durations and count of              occurrences.
  4. The script saves the updated workbook.
