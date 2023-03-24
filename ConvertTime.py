import datetime

seconds_since_1900 = int(input("Enter the number of seconds since January 1, 1900: "))

# Convert seconds to days and seconds
days, seconds = divmod(seconds_since_1900, 86400)

# Add days to January 1, 1900 to get the datetime object
base_date = datetime.datetime(1900, 1, 1)
date = base_date + datetime.timedelta(days=days, seconds=seconds)

# Output the result
print("The date and time corresponding to", seconds_since_1900, "seconds since January 1, 1900 is:")
print(date.strftime("%Y-%m-%d %H:%M:%S"))
