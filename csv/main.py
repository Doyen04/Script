import pandas as pd

file_path = './dateofbirth.xlsx'

data = pd.read_excel(file_path)
print(data.head())
data['DateofBirth'] = pd.to_datetime(data['DateofBirth'], format='%d %B,%Y').dt.strftime('%Y-%m-%d %H:%M:%S.%f')
data['DateofBirth'] = data['DateofBirth'].str[:-3]
print(data.head())
text_file_path = './output.txt'

with open(text_file_path, 'a') as file:
    for index, row in data.iterrows():
        text = f"update PolicyMaster set DateofBirth='{row['DateofBirth']}' where PolicyNo='{row['PolicyNo']}'\n"
        file.write(text)
        # if index == 10: break