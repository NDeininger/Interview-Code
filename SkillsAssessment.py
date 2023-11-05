import pandas as pd
import os

#Set directory to where this Python script is
script_directory = os.path.dirname(os.path.abspath(__file__))
os.chdir(script_directory)

#Load PRODUCTION_HISTORY data
prodHistData = pd.read_excel('data.xlsx', 'PRODUCTION_HISTORY')
#Load MACHINE_PROD_STANDARDS data
machineStandardsData = pd.read_excel('data.xlsx', 'MACHINE_PROD_STANDARDS')

#Drop downtime entries
prodHistData = prodHistData[~(prodHistData['customer #'].str.contains('downtime') & prodHistData['product_id'].astype(str).str.contains('downtime'))]

#Merge machine standards with prodHistData
prodHistData = prodHistData.merge(machineStandardsData, on=['machine_id', 'product_id'], how='left')

#Calculate expected job completion time 
prodHistData['expected_job_completion_time(hrs)'] = prodHistData['order pounds for calcs'] / prodHistData['standard_lbs_per_hour']

#Calculate the average expected job completion and the actual job completion time
average_data = prodHistData.groupby('machine_id').agg({
    'job_completion_time(hrs)': 'mean',
    'expected_job_completion_time(hrs)': 'mean'
}).reset_index()


with pd.ExcelWriter('Analysis.xlsx', engine='openpyxl') as writer:
    # Write 'prodHistData' to the 'Production_History_Data' sheet
    prodHistData.to_excel(writer, sheet_name='Production_History_Data', index=False)
    # Write 'average_data' to the 'Average_Data' sheet
    average_data.to_excel(writer, sheet_name='Average_Data', index=False)

