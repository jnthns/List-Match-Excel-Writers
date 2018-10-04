### READ 'NO EXPANSION' MATCH FILE AND WRITE TO EXCEL FILE WITH 3 WORKSHEETS

### Paste the match output file into the folder that contains this script.
### Paste the match filename into the variable 'f' and the desired excel filename 
### into the variable 'f2' below.

import pandas as pd  
import os

### READ CSV 
f = 'Comdata Ads List 2 - Q4 2018 - 10_2_18_2018-10-02_132785_output.csv'

# Create dataframe for entire match file to obtain number of columns(length).
matchfile_df = pd.read_csv(f)
length = len(matchfile_df.columns)

# Select column ranges to include:
range1 = [0,1,2,7,8,9,10,14,19,20,21,28]
# brand expansion range below
# range1 = [0,1,2,4,7,8,12,13,17,22,23,24,30]

# Range for all custom attribute columns
range2 = list(range(35,length))

# List of columns to include
cols = range1 + range2

# Create dataframe with selected columns
df = pd.read_csv(f, skiprows=None, usecols=cols)

#-----------------------------------------------------------------------------------------------------------------------

### FILTER DATA

# Remove rows with 'active' == 'False' from reachable and not reachable

# Create 'reachable' dataframe for 'MEI' >= 20 and 'active' == True
reachable_df = df[(df['MEI'] >= 20) & (df['active'] == True)]
reachable_df = reachable_df.sort_values(by=['MEI'], ascending=False)


# Create 'Not_Reachable' dataframe for 'MEI' < 20 and 'active' == True
not_reachable_df = df[(df['MEI'] < 20) & (df['active'] == True)]
not_reachable_df = not_reachable_df.sort_values(by=['MEI'], ascending=False)

# Create 'No_Match' dataframe for rows that require manual scrubbing
# These rows also have 'active' == False and do not contain any duplicates
no_match_df = df[(df['active'] == False) & (df['Match Type'] != 'duplicate input') & (df['Match Type'] != 'duplicate match')]

# Create 'Duplicates' dataframe of rows with duplicate input or duplicate match
duplicates_df = df[((df['Match Type'] == 'duplicate input') | (df['Match Type'] == 'duplicate match'))]

# Drop unwanted columns in each dataframe
cols = ['active', 'Match Result', 'Country Name']
reachable_df = reachable_df.drop(columns = cols)
not_reachable_df = not_reachable_df.drop(columns = cols)
no_match_df = no_match_df.drop(columns = cols)
duplicates_df = duplicates_df.drop(columns = cols)

#----------------------------------------------------------------------------------------------------------------------

### WRITE XLSX

# Excel filename
f2 = 'Comdata Ads List 2 - Q4 2018 - 10_2_18_2018-10-02_132785_output.xlsx'

os.chdir('/Users/jshek/Desktop/completedxls')

writer = pd.ExcelWriter(f2, engine = 'xlsxwriter')
reachable_df.to_excel(writer, sheet_name = 'Reachable', index = False)
not_reachable_df.to_excel(writer, sheet_name = 'Not Reachable', index = False)
no_match_df.to_excel(writer, sheet_name = 'No Match', index = False)
duplicates_df.to_excel(writer, sheet_name = 'Duplicates', index = False)

writer.save()
