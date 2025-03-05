import pandas as pd
import requests
from io import StringIO
from tkinter import filedialog
from time import sleep

 
# Fetch the CSV data from the API
url = 'https://data-api.ecb.europa.eu/service/data/EXR/D.NOK.EUR.SP00.A?format=csvdata' 
response = requests.get(url)
csv_data = StringIO(response.text)
 
# Load the data into a Pandas DataFrame
df = pd.read_csv(csv_data)
 
# Convert 'TIME_PERIOD' to datetime format
df['TIME_PERIOD'] = pd.to_datetime(df['TIME_PERIOD'])
df['OBS_VALUE'] = pd.to_numeric(df['OBS_VALUE'], errors='coerce')
 
# Drop rows with missing values in 'TIME_PERIOD' or 'OBS_VALUE'
df.dropna(subset=['TIME_PERIOD', 'OBS_VALUE'], inplace=True)
 
# Function to calculate average for a given time period
def calculate_average(start_date, end_date):
    start_date = pd.to_datetime(start_date, format='%Y-%m-%d')
    end_date = pd.to_datetime(end_date, format='%Y-%m-%d')
    
    # Filter data within the specified date range
    filtered_data = df[(df['TIME_PERIOD'] >= start_date) & (df['TIME_PERIOD'] <= end_date)]
    
    # Calculate the average of 'OBS_VALUE' in the filtered data
    average_value = filtered_data['OBS_VALUE'].mean()
    
    return average_value

# Load the Excel file containing student data
filename = filedialog.askopenfilename()
student_data = pd.read_excel(filename)

# Prepare a new column for the averages
student_data['Snittkurs'] = student_data.apply(lambda row: calculate_average(row['Stay from'], row['Stay to']), axis=1)


#Calculate the remaining grant based on previous payments
flat_first_rate = 11.7
def calculate_remaining(total, first, rate):
    return total*rate - first*flat_first_rate


student_data['Siste utbetaling'] = student_data.apply(lambda row: 
                                                      calculate_remaining(row['Total grant'], row['First payment'], row['Snittkurs']), axis=1)


# Save the output to a new Excel file
student_data.to_excel(filename, index=False)

print("Success!")
sleep(3)
