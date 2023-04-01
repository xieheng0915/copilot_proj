# write a funciton to return all excel files in a directory

import os


def get_excel_files(path):
    files = os.listdir(path)
    excel_files = []
    for file in files:
        if file.endswith('.xlsx'):
            excel_files.append(file)
    return excel_files

path = "/Users/xieheng/gitspace/python3space/copilot_with_py/data"

# Get all excel files in the directory
excel_files = get_excel_files(path)

# print the excel files 
print(excel_files)

# return dataframe from excel path
import pandas as pd

def get_df_from_excel(path, sheet_name):
    df = pd.read_excel(path, sheet_name=sheet_name)
    return df

# create a new folder "data_new" to store the new excel files
new_path = "/Users/xieheng/gitspace/python3space/copilot_with_py/data_new"
if not os.path.exists(new_path):
    os.makedirs(new_path)

# get each dataframe from each excel file

for file in excel_files:
    df = get_df_from_excel(path + "/" + file, sheet_name="Sheet1")
    # print(df)
    # replace each cell to 0 if the value is not a number
    df = df.applymap(lambda x: 0 if not isinstance(x, (int, float)) else x)
    # Sum each row
    df["Sum"] = df.sum(axis=1)
    # save the new dataframe to a new excel file by the same name add "_new" to the end
    df.to_excel(os.path.join(new_path, os.path.basename(file).split(".")[0] + "_new.xlsx"), index=False)
    
    print(df)