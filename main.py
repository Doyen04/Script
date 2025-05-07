import pandas as pd
from pandas.tseries.offsets import BDay
import random
from datetime import datetime, timedelta
# Load the entire Excel file
excel_file = './BACKUP-RESTORE.xlsx'

# Read all sheets into a dictionary of DataFrames
all_sheets = pd.read_excel(excel_file, sheet_name=None)

# Access individual sheets by their names
sheet1 = all_sheets['BACKUP']
sheet2 = all_sheets['RESTORE']

# Generate sequential working days for 2025 with random times
def generate_working_days_with_time(start_date, num_days):
    dates = pd.date_range(start=start_date, periods=num_days, freq='B')  # Use 'B' for business days
    times = [timedelta(hours=random.randint(0, 23), minutes=random.randint(0, 59), seconds=random.randint(0, 59)) for _ in range(num_days)]
    return [date + time for date, time in zip(dates, times)]

# Generate new working day dates for the columns
num_backup_rows = len(sheet1)
num_restore_rows = len(sheet2)

# Generate working days with times for backup and restore
backup_start_dates = generate_working_days_with_time('2025-01-01', num_backup_rows)
backup_finish_dates = generate_working_days_with_time('2025-01-01', num_backup_rows)

if len(sheet2) < len(sheet1):
    # Calculate how many rows to duplicate
    rows_to_add = len(sheet1) - len(sheet2)
    
    # Duplicate rows from sheet2
    duplicated_rows = sheet2.sample(n=rows_to_add, replace=True).reset_index(drop=True)
    
    # Append duplicated rows to sheet2
    sheet2 = pd.concat([sheet2, duplicated_rows], ignore_index=True)

# Generate new working day dates for the extended sheet2
restore_dates = generate_working_days_with_time('2025-01-01', len(sheet2))

# Assign the generated dates to the Restore Date column
sheet2['Restore Date'] = restore_dates
# Assign the generated dates to the respective columns
sheet1['backup_start_date'] = backup_start_dates
sheet1['backup_finish_date'] = backup_finish_dates

import re

# Function to update the date in the string to match the backup_start_date
def update_date_in_string(original_string, new_date):
    # Check if the string contains a date in the format YYYY_MM_DD
    if re.search(r'\d{4}_\d{2}_\d{2}', original_string):
        # Replace the date with the new date
        updated_string = re.sub(r'\d{4}_\d{2}_\d{2}', new_date.strftime('%Y_%m_%d'), original_string)
        return updated_string
    else:
        # If no date is found, return the original string
        return original_string

# Update the physical_device_name and backupset_name columns to match backup_start_date
sheet1['physical_device_name'] = sheet1.apply(
    lambda row: update_date_in_string(row['physical_device_name'], row['backup_start_date']), axis=1
)

sheet1['backupset_name'] = sheet1.apply(
    lambda row: update_date_in_string(row['backupset_name'], row['backup_start_date']), axis=1
)


with open('backup.txt', 'a') as file:
    for index, sheet in sheet1.iterrows():
        for k in range(4):
            start_offset = timedelta(hours=random.randint(0, 2), minutes=random.randint(0, 59), seconds=random.randint(0, 59))
            finish_offset = start_offset + timedelta(minutes=random.randint(1, 120))  # Ensure finish is always later

            # Apply offsets to the original backup_start_date and backup_finish_date
            start_date_with_offset = sheet['backup_start_date'] + start_offset
            finish_date_with_offset = sheet['backup_start_date'] + finish_offset

            file.write(f"""
            INSERT INTO [dbo].[GIBS_Backup]
                ([Server]
                ,[database_name]
                ,[backup_start_date]
                ,[backup_finish_date]
                ,[expiration_date]
                ,[backup_type]
                ,[backup_size]
                ,[logical_device_name]
                ,[physical_device_name]
                ,[backupset_name]
                ,[description])
            VALUES
                ('NSIA-NG-GIBSSQL'
                ,'{sheet['database_name']}'
                ,'{start_date_with_offset.strftime('%Y-%m-%d %H:%M:%S')}'
                ,'{finish_date_with_offset.strftime('%Y-%m-%d %H:%M:%S')}'
                ,'{'' if pd.isna(sheet['expiration_date']) else sheet['expiration_date']}'
                ,'Database'
                ,'{sheet['backup_size']}'
                ,'{'' if pd.isna(sheet['logical_device_name']) else sheet['logical_device_name']}'
                ,'{sheet['physical_device_name']}'
                ,'{sheet['backupset_name']}'
                ,'{'' if pd.isna(sheet['description']) else sheet['description']}');
                GO
        """)
    
with open('restore.txt', 'a') as file:
    for index, sheet in sheet2.iterrows():
        for k in range(4):
            restore_date_with_offset = sheet['Restore Date'] + timedelta(
                hours=random.randint(0, 1), minutes=random.randint(0, 59), seconds=random.randint(0, 59)
            )
            file.write(f"""
            INSERT INTO [dbo].[GIBS_Restore]
                ([DatabaseName]
                ,[RestoreType]
                ,[RestoreDate]
                ,[Source]
                ,[RestoreFile]
                ,[RestoredBy])
            VALUES
                ('Gibs5db_NSIA'
                ,'Database'
                ,'{restore_date_with_offset.strftime('%Y-%m-%d %H:%M:%S')}'
                ,'{sheet['Source']}'
                ,'{sheet['Restore File']}'
                ,'NSIAINSURANCE\\NSIA-NG-GIBSQL0$');
            GO
            """)