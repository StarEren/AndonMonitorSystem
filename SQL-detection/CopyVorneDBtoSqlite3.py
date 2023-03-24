import sqlite3
import requests

# Set up the API endpoint and parameters
url = 'http://10.10.12.103/sql-request.do'

params = {
    'response_type': 'application/json',
    'sql_statement': 'SELECT * FROM timeline_stream LIMIT 10;'
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

# Create a new SQLite database and table
conn = sqlite3.connect('timeline.db')
c = conn.cursor()


c.execute('''CREATE TABLE IF NOT EXISTS timeline
             (col1 KEY,
              col2 INTEGER,
              col3 TEXT,
              col4 TEXT,
              col5 TEXT,
              col6 TEXT,
              col7 TEXT,
              col8 TEXT,
              col9 TEXT)''')

# Insert the rows into the SQLite database
for row in rows:
    c.execute("INSERT INTO timeline (col1, col2, col3, col4, col5, col6, col7, col8, col9) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)",
                (row[0], row[1], row[2], row[3], row[4], row[5], row[6], row[7], row[8]))
# Commit the changes and close the connection
conn.commit()
conn.close()