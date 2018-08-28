import pandas as pd
import numpy as np

# Read data file 
df = pd.read_excel("..//data//test.xlsx")

#Add hour column 
df['hours'] = df['Time'].dt.hour

#Fill NaN with 0
df = df.fillna(value=0,axis='columns')

#Constants
avg_vehicle_mil = 1.0  #1kwh/1km
max_distance = 3.74    #3.74 KM
avg_commute_time = 14  #14min/hr

'''
Peak Hours detection
• 0.95 for congested condition
• 0.92 for urban areas
• 0.88 for rural areas
• 0.85 for minor street inflows and outflows
• 0.90 for minor arterial
refrence: https://nptel.ac.in/courses/105101008/downloads/cete_05.pdf
'''

peak_hour_threshold = 0.92

#Calculate PCU 
df['PCU'] = (df['Car'] * 1) + (df['HCV'] * 3.5) + (df['LCV'] * 2.2) + (df['3W'] * 0.8) + (df['2W'] * 0.5) + (df['BICYCLE'] * 0.2)

# Calculate sum PCU/HR
df_hour = df.groupby(['hours'])['PCU'].sum()

# Calculate Max PCU
df_max = df.groupby(['hours'])['PCU'].max()

# Creating New Dataframe
df_new = pd.DataFrame()

# Adding hours
df_new['hours'] = df_hour.keys()

# Sum PCUHR = Sum(PCU)
df_new['PCUHR'] = df_hour

# PHF = PCUHR / (4 * Max(PCU))
df_new['PHF'] = (df_hour/(4*df_max))

# Vehicle_Flow_Hr = PCUHR / PHF
df_new['Vehicle_Flow_Hr'] = (df_new['PCUHR']/df_new['PHF'])

# Power_load_region_Kwh = Vehicle_Flow_Hr * avg_vehicle_mil * max_distance
df_new['Power_load_region_Kwh'] = (df_new['Vehicle_Flow_Hr'] * avg_vehicle_mil * max_distance)

# Actual_load_region_Kwh = Power_load_region_Kwh + 15%
df_new['Actual_load_region_Kwh'] = (df_new['Power_load_region_Kwh'] * 1.15)

# ABSV_required = Actual_load_region_Kwh / 3000
df_new['ABSV_required'] = (df_new['Actual_load_region_Kwh'] / 3000)

# Actual_ABSV_required = ceil(ABSV_required)
df_new['Actual_ABSV_required'] = (np.ceil(df_new['ABSV_required']))

# Avg_Access_time_per_vehicle = max_distance / (avg_commute_time / 60)
df_new['Avg_Access_time_per_vehicle'] = (max_distance / (avg_commute_time / 60))

# Find peak hours if PHF > peak_hour_threshold
df_new['is_peak'] = np.where(df_new['PHF']>=peak_hour_threshold, 'yes', 'no')

writer = pd.ExcelWriter('output.xlsx')
df.to_excel(writer,'15Min')
df_new.to_excel(writer,'Hours')
writer.save()
