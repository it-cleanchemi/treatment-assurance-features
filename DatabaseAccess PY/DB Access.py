import psycopg2
import csv
from datetime import datetime, timedelta
import pytz

# Connect to the PostgreSQL database
conn = psycopg2.connect(
    dbname="cleanchemidata",
    user="cc_reader",
    password="3G8oNnDRsNfi",
    host="35.202.185.70",
    port="5432"
)

# Open a cursor to perform database operations
cur = conn.cursor()

# Get the timezone for the timestamps stored in your database
# Replace 'YourDatabaseTimezone' with the actual timezone used in your database
database_timezone = pytz.timezone('Etc/GMT-5')

# Calculate the current time in the database timezone
current_time = datetime.now(database_timezone)

# Calculate the timestamp representing 24 hours ago in the database timezone
twenty_four_hours_ago = current_time - timedelta(hours=24)

# Execute the query to retrieve data from the last 24 hours
cur.execute("SELECT * FROM cc_systemdata WHERE time_stamp >= %s ORDER BY time_stamp DESC", (twenty_four_hours_ago,))

# Fetch all rows from the result set
rows = cur.fetchall()

# Define the path to save the CSV file
csv_file = 'C:/Users/Vlad/Test.csv'

# Write the query results to a CSV file
with open(csv_file, 'w', newline='') as file:
    writer = csv.writer(file)
    writer.writerow([desc[0] for desc in cur.description])  # Write column headers
    writer.writerows(rows)  # Write rows

# Close the cursor and connection
cur.close()
conn.close()

print("Data exported successfully to", csv_file)
