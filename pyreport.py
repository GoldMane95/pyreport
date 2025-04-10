import pandas as pd
import os
from tqdm import tqdm

# Above imports pandas and os for working with excel and operating system
# tqdm shows progress of file in terminal

# Variables to house folder paths
folder_path1 = 'base'
folder_path2 = 'extra'
folder_path3 = 'output'

# Read contents within designated folder paths
all_files_base = [f for f in os.listdir(folder_path1) if f.endswith('.xlsx')]
all_files_extra = [f for f in os.listdir(folder_path2) if f.endswith('.xlsx')]


# Empty lists to hold content of files for each folder
dataframes1 = []
dataframes2 = []

# Extract data from files within folders, add to dataframe list
for file in tqdm(all_files_base, desc="Processing base folder"):
    file_path = os.path.join(folder_path1, file)
    df = pd.read_excel(file_path, engine='openpyxl')
    dataframes1.append(df)

for file in tqdm(all_files_extra, desc="Processing extra folder"):
    file_path = os.path.join(folder_path2, file)
    df = pd.read_excel(file_path, engine='openpyxl', skiprows=7)
    dataframes2.append(df)

# Combine the data of each file within dataframe
combined_df1 = pd.concat(dataframes1, ignore_index=True)
combined_df2 = pd.concat(dataframes2, ignore_index=True)

# Print the column names of combined_df2 to verify
print("Columns in combined_df1:", combined_df1.columns)
print("Columns in combined_df2:", combined_df2.columns)

# Remove duplicate rows based on columns A, B, and C in combined_df2
combined_df2.drop_duplicates(subset=['Unnamed: 0', 'Unnamed: 1', 'Unnamed: 2'], inplace=True)

# Verify that columns D, F, and G of each row in combined_df1 match with columns A, B, and C of combined_df2 respectively
# If they match then append columns D, E, and G of combined_df2 to the end of the matching row in combined_df1
for index2, row2 in tqdm(combined_df2.iterrows(), desc="Verifying and combining data", total=combined_df2.shape[0]):
    for index1, row1 in combined_df1.iterrows():
        if (row2['Unnamed: 0'] == row1['Home Department Description']) and (row2['Unnamed: 1'] == row1['Last Name']) and (row2['Unnamed: 2'] == row1['First Name']):
            combined_df1.at[index1, 'Curriculum'] = row2['Unnamed: 3']
            combined_df1.at[index1, 'Status'] = row2['Unnamed: 4']
            combined_df1.at[index1, 'Due Date'] = row2['Unnamed: 6']

# Create new file in output folder
output_path = os.path.join(folder_path3, 'output.xlsx')
combined_df1.to_excel(output_path, index=False)


print(f"All Excel files in the folder have been combined into '{output_path}'.")