import pandas as pd

dataFrame = pd.read_excel("C:\\Users\\arhsathy\\Desktop\\Automation.xlsx")
print(dataFrame.index)

for i in dataFrame.index:
    # print(dataFrame.loc[i])
    content='\nWe are learing python'
    with open(f"{dataFrame.loc[i]['Name']}",'w') as file:
        file.write(f"Hi {dataFrame.loc[i]['Name']},\n"+content)
