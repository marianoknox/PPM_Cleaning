# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""
# import dependencies
import pandas as pd
import numpy as np

# read files
## main file
df = pd.read_excel("./PPM2.xlsx")
row = len(df)
col = len(df.columns)
print("No. of Rows:" + str(row))
print("No. of Columns:" + str(col))
df.head()

## reference files
### Location References
dfLocation = pd.read_excel("./Ref.xlsx", "Location")
dfLocation.head()

### Department References
dfDepartment = pd.read_excel("./Ref.xlsx", "Department")
dfDepartment.head()

### Frequency References
dfFrequency = pd.read_excel("./Ref.xlsx", "Frequency")
dfFrequency.head()

### System Refertences
dfSystem = pd.read_excel("./Ref.xlsx", "System")
dfSystem.head()

# Drop unnecessary columns
try :
    del df['Flag']
except : 
    print('Column Deleted')
    
row = len(df)
col = len(df.columns)
print("No. of Rows:" + str(row))
print("No. of Columns:" + str(col))

# Creat additional comuns
## Adding Status 2 and 3
def getStatus(txtFrom) :
    if "Scheduled" in txtFrom :
        return "Open"
    elif "In Progress" in txtFrom :
        return "Open"
    elif "New" in txtFrom :
        return "Open"
    elif "On Hold" in txtFrom :
        return "On Hold"
    elif "Closed" in txtFrom :
        return "Closed"
    elif "Not Done" in txtFrom :
        return "Closed"
    elif "Rejected" in txtFrom :
        return "Closed"
    else : return "NA"
    
def getStatus2(txtFrom) :
    if "On Hold" in txtFrom :
        return "Open"
    else : return txtFrom

df['Status 2'] = df['Status'].apply(lambda x: getStatus(x))
df['Status 3'] = df['Status 2'].apply(lambda x: getStatus2(x))

row = len(df)
col = len(df.columns)
print("No. of Rows:" + str(row))
print("No. of Columns:" + str(col))

df[['WO','Status','Status 2','Status 3']].head()

