# Winter Workterm - Vorne (Andon) Database Project


## This is a work term project using camera to develop a system that pulls data from Vorne Andon system, convert to a readable format, and export to an excel file that is emailed out periodically.   

### Our process:

- Setting up 3 different database connections using Vorne's API.
- Pulling data using LESQL query parameters.
- Converting the data from integer code to a string.
- Placing data into excel file.
- Automatic email of excel file that is sent out periodically.

### Notes:

- 3rd column is time in seconds, starting from Jan 1, 1990. Can be converted to current timestamp with python program that I've made.
- 7th column is the error codes. CSV file of error code definitions be found in repo.

### Requirements:

- [Python 3.9 to 3.11.1](https://www.python.org/downloads/release/python-3111/) is required.
- [NodeJS 16](https://nodejs.org/en/download/) is required.

### Local Development:

```
# Clone the repo
$ git clone https://github.com/StarEren/VorneDB
# Move into directories
$ cd <directory-name>
# Install the requirements to run the program
$ pip install <packages-name>
# Start the program and have fun!!!
$ python main.py
```
