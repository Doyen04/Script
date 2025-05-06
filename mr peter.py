import pandas as pd

# Load the entire Excel file
excel_file = './BACKUP-RESTORE.xlsx'

# Read all sheets into a dictionary of DataFrames
all_sheets = pd.read_excel(excel_file, sheet_name=None)

# Access individual sheets by their names
sheet1 = all_sheets['BACKUP']
sheet2 = all_sheets['RESTORE']

# Safely update the year to 2025 for backup_start_date and backup_finish_date
def safe_replace_year(date, year):
    try:
        return date.replace(year=year)
    except ValueError:
        # Handle invalid dates (e.g., February 29 in non-leap years)
        return date + pd.offsets.DateOffset(year=year)

sheet1['backup_start_date'] = pd.to_datetime(sheet1['backup_start_date']).apply(lambda x: safe_replace_year(x, 2025))
sheet1['backup_finish_date'] = pd.to_datetime(sheet1['backup_finish_date']).apply(lambda x: safe_replace_year(x, 2025))
sheet2['Restore Date'] = pd.to_datetime(sheet2['Restore Date']).apply(lambda x: safe_replace_year(x, 2025))
# Print the updated columns

with open('backup.txt', 'a') as file:
    for index, sheet in sheet1.iterrows():
        file.write(f"""
        INSERT INTO [dbo].[GIBS_Backup]
            SELECT
            [Server]
            ,[database_name]
            ,[backup_start_date]
            ,[backup_finish_date]
            ,[expiration_date]
            ,[backup_type]
            ,[backup_size]
            ,[logical_device_name]
            ,[physical_device_name]
            ,[backupset_name]
            ,[description]
        VALUES
            'NSIA-NG-GIBSSQL'
            ,'{sheet['database_name']}'
            ,'{sheet['backup_start_date']}'
            ,'{sheet['backup_finish_date']}'
            ,'{sheet['expiration_date']}'
            ,'Database'
            ,'{sheet['backup_size']}'
            ,'{sheet['logical_device_name']}'
            ,'{sheet['physical_device_name']}'
            ,'{sheet['backupset_name']}'
            ,'{sheet['description']}'
            GO
    """)
    
with open('restore.txt', 'a') as file:
    for index, sheet in sheet2.iterrows():
        file.write(f"""
        INSERT INTO [dbo].[GIBS_Restore]
            SELECT
            [DatabaseName]
            ,[RestoreType]
            ,[RestoreDate]
            ,[Source]
            ,[RestoreFile]
            ,[RestoredBy]
        VALUES
            'Gibs5db_NSIA'
            ,'Database'
            ,'{sheet['Restore Date']}'
            ,'{sheet['Source']}'
            ,'{sheet['Restore File']}'
            ,'NSIAINSURANCE\\NSIA-NG-GIBSQL0$'
        GO
        """)