import pandas as pd
from tkinter import filedialog
import os

getpass=os.getlogin()
sys_username=getpass

input_file=filedialog.askopenfilename(title="Select the Input file",filetypes=(("Excel Files", "*.xlsx"), ("CSV Files", "*.csv"), ("All Files", "*.*")))
df=pd.read_excel(input_file)

row_limit=3000

row_count=len(df)
print(row_count)
num_files=(row_count-1)//row_limit+1
print(num_files)


for i in range(num_files):
    start=i*row_limit
    end=min((i+1)*row_limit,row_count)
    df_base=df.iloc[start:end]
    output_file="C:\\Users\\"+str(sys_username)+f"\\Desktop\\Splitting Output files\\Excel_Splitted_file__{i+1}.xlsx"
    df_base.to_excel(output_file,index=False)

print('Executed Successfully.....')

