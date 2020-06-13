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

## Adding Line and Station Columns
def getInput(listOfRef, txtFrom) :
    #listOfRef = dfLocation items
    #txtFrom = df current index

    for curRef in listOfRef :
        if curRef in txtFrom :
            return curRef
    return "not available"

df['Code'] = df['Location'].apply(lambda x: getInput(dfLocation['Code'], x))
dfWithLocation = pd.merge(df, dfLocation)

row = len(df)
col = len(df.columns)
print("Original No. of Rows:" + str(row))
print("Original No. of Columns:" + str(col))
print("-------------------------------------")
rowN = len(dfWithLocation)
colN = len(dfWithLocation.columns)
print("New No. of Rows:" + str(rowN))
print("New No. of Columns:" + str(colN))
print("-------------------------------------")
rowD = rowN - row
colD = colN - col
print("Difference in No. of Rows:" + str(rowD))
print("Difference in No. of Columns:" + str(colD))

dfWithLocation[['WO','Location','Code','Line','Station']].head()

#test = dfWithLocation.loc[dfWithLocation['']=="mech"]
#test[['WO','Assigned To','Department','Department 2']].head(100)

## Adding Frequency Column
#Fill Empty Task Field based on Frequency
dfWithLocation['Task/Work Subject Group'] = dfWithLocation['Task/Work Subject Group'].fillna("Annual")

dfWithLocation['Task/Work Subject Group'] = dfWithLocation['Task/Work Subject Group'].str.lower()

def getInput(listOfRef, txtFrom) :
    for curRef in listOfRef :
        if curRef in txtFrom :
            return curRef.lower()
    return "not available"


dfWithLocation['FreqRef'] = dfWithLocation['Task/Work Subject Group'].apply(lambda x: getInput(dfFrequency['FreqRef'], x))
dfWithLocation = pd.merge(dfWithLocation, dfFrequency)

row = len(df)
col = len(df.columns)
print("Original No. of Rows:" + str(row))
print("Original No. of Columns:" + str(col))
print("-------------------------------------")
rowN = len(dfWithLocation)
colN = len(dfWithLocation.columns)
print("New No. of Rows:" + str(rowN))
print("New No. of Columns:" + str(colN))
print("-------------------------------------")
rowD = rowN - row
colD = colN - col
print("Difference in No. of Rows:" + str(rowD))
print("Difference in No. of Columns:" + str(colD))

#Drop Reference
try :
    del dfWithLocation['FreqRef']
except : 
    print('*****Ref Column Deleted******')

dfWithLocation[['WO','Task/Work Subject Group','Frequency']].head()

## Adding system column
def getInput(listOfRef, txtFrom) :
    for curRef in listOfRef :
        if curRef in txtFrom :
            return curRef.lower()
    return "not available"
    

dfWithLocation['SysRef'] = dfWithLocation['Task/Work Subject Group'].apply(lambda x: getInput(dfSystem['SysRef'], x))
dfWithLocation = pd.merge(dfWithLocation, dfSystem)

row = len(df)
col = len(df.columns)
print("Original No. of Rows:" + str(row))
print("Original No. of Columns:" + str(col))
print("-------------------------------------")
rowN = len(dfWithLocation)
colN = len(dfWithLocation.columns)
print("New No. of Rows:" + str(rowN))
print("New No. of Columns:" + str(colN))
print("-------------------------------------")
rowD = rowN - row
colD = colN - col
print("Difference in No. of Rows:" + str(rowD))
print("Difference in No. of Columns:" + str(colD))

#Drop Reference
try :
    del dfWithLocation['SysRef']
except : 
    print('*****Ref Column Deleted******')

dfWithLocation[['WO','Task/Work Subject Group','System']].head(3000)

# Checking fo Null and Supplying Data
nullRows = dfWithLocation['Department'].isna().sum()
print("Department Initial Null count: " + str(nullRows))
nullRows2 = dfWithLocation['Assigned To'].isna().sum()
print("Assigned To Initial Null count: " + str(nullRows2))

#Fill empty Assigned To Field
#dfWithLocation['Department'].fillna(dfWithLocation['Assigned To'], inplace=True) -- For CM

dfWithLocation['Assigned To'] = dfWithLocation['Assigned To'].str.lower()

def checkDept(txtFrom, listOfRef):
    for curRef in listOfRef :
        if curRef in txtFrom :
            return curRef.lower()
    return "not available"

        
dfWithLocation['Department'] = dfWithLocation['Assigned To'] \
.apply(lambda x: checkDept(x, dfDepartment['Department']))

dfWithLocation = pd.merge(dfWithLocation, dfDepartment)

#dfWithLocation['Department'] = dfWithLocation['Department'].apply(lambda x: dfDepartment.loc[x['Department'] == dfDepartment['Department']])

#dfWithLocation['Department 2'] = dfDepartment.loc[dfDepartment['Department'] == dfWithLocation['Department'],['Department 2']]

nullRows2 = dfWithLocation['Department'].isna().sum()
print("After supplying Null count: " + str(nullRows2))

#dfWithLocation[['WO','Assigned To','Department','Department 2']].head(100)
#dfWithLocation.loc[dfWithLocation['Department']=="mech"]

## Department cross checking
test = dfWithLocation.loc[dfWithLocation['Department']=="mech"]
test[['WO','Assigned To','Department','Department 2']].head(100)

## Converting Date format to m/d/yyyy
# try :
#     dfWithLocation['Planned Start'] = pd.to_datetime(dfWithLocation['Planned Start'], format='%Y%m%d').dt.strftime('%m/%d/%Y')
#     dfWithLocation['Raised Date'] = pd.to_datetime(dfWithLocation['Raised Date'], format='%Y%m%d').dt.strftime('%m/%d/%Y')
#     dfWithLocation['Actual Completion Date / Time'] = pd.to_datetime(dfWithLocation['Actual Completion Date / Time'], \
#                                                                      format='%Y%m%d').dt.strftime('%m/%d/%Y')
# except :
#     print("***Formated Already")
    
# dfWithLocation['Planned Start'] = dfWithLocation['Planned Start'].astype('datetime64[ns]')
# dfWithLocation['Raised Date'] = dfWithLocation['Raised Date'].astype('datetime64[ns]')
# dfWithLocation['Actual Completion Date / Time'] = dfWithLocation['Actual Completion Date / Time'].astype('datetime64[ns]')

# dfWithLocation[['WO','Planned Start','Raised Date','Actual Completion Date / Time']].head(20)


# Re-arranging Orders and Export
## Column Order
rowN = len(dfWithLocation)
colN = len(dfWithLocation.columns)
print("New No. of Rows:" + str(rowN))
print("New No. of Columns:" + str(colN))

row = len(dfWithLocation)
#dfComplete = dfWithLocation[['WO','Code','Line','Station','Status 2','Description', 'Location', 'Assigned To','Department 2', \
                             #'Planned Start','Raised Date','Actual Completion Date / Time']][0:row]
    

dfComplete = dfWithLocation
dfComplete.head()

## Export to Excel
## Compare to Original File
row = len(df)
col = len(df.columns)
print("Original No. of Rows:" + str(row))
print("Original No. of Columns:" + str(col))
print("-------------------------------------")
rowN = len(dfWithLocation)
colN = len(dfWithLocation.columns)
print("New No. of Rows:" + str(rowN))
print("New No. of Columns:" + str(colN))
print("-------------------------------------")
rowD = rowN - row
colD = colN - col
print("Difference in No. of Rows:" + str(rowD))
print("Difference in No. of Columns:" + str(colD))

## Export
#dfComplete.to_excel("PPM output.xlsx")

