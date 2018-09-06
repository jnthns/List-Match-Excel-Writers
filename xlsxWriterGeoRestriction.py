### WRITE OUTPUT FILE TO EXCEL FILE WITH GEO RESTRICTION(i.e. US and CA only)
### IF MATCH IS NOT IN DESIRED REGION, MOVE TO 'NO MATCH' TAB

### Paste the match output file into the folder that contains this script.
### Paste the match filename into the variable 'f' and the desired excel filename 
### into the variable 'f2' below.

import pandas as pd
import sys  

# Changed default string encoding from 'ascii' to 'utf8'
reload(sys)  
sys.setdefaultencoding('utf8')

### READ CSV 

# Match filename
#f = "DME_Output_Duplicates_Column.csv"
f = 'Veritas AMS Ent East_2018-07-26_127813_output (1).csv'

# Create dataframe for entire match file to obtain number of columns(length).
matchfile_df = pd.read_csv(f, skiprows=None)
length = len(matchfile_df.columns)

# Select column ranges to include:
# Source Name, Match Result, Match Type, Best SID, Account, Country Code, MEI
#range1 = [0,1,2,10,21,4,7,8,9,19,20,28, 31]
range1 = [0,1,2,4,7,8,9,10,19,20,21,28, 31]

# Range for all custom attribute columns
range2 = range(35,length)

# List of columns to include
cols = range1 + range2

# Create dataframe with selected columns
df = pd.read_csv(f, skiprows=None, usecols=cols)


#-----------------------------------------------------------------------------------------------------------------------

### FILTER DATA ON MEI AND GEO RESTRICTION

# Remove rows with 'active' == 'False' from reachable and not reachable

# Create 'reachable' dataframe for 'MEI' >= 25 and 'active' == True
#reachable_df = df[(df['MEI'] >= 20) & (df['active'] == True) & (df['duplicate'] == False) & (df['Industry'] != 'Education') & (df['Industry'] != 'Government')]
reachable_df = df[(df['MEI'] >= 20) & (df['active'] == True) \
	& (df['Country Code'] == 'US') | (df['Country Code'] == 'CA')]

# Create 'Not_Reachable' dataframe for 'MEI' < 25 and 'active' == True
#not_reachable_df = df[(df['MEI'] < 20) & (df['active'] == True) & (df['duplicate'] == False) & (df['Industry'] != 'Education') & (df['Industry'] != 'Government')]
not_reachable_df = df[(df['MEI'] < 20) & (df['active'] == True) \
	& (df['Country Code'] == 'US') | (df['Country Code'] == 'CA')]

# Create 'No_Match' dataframe for rows that require manual scrubbing
# These rows also have 'active' == False and do not contain any duplicates
no_match_df = df[(df['active'] == False) & (df['Match Type'] != 'duplicate input') & (df['Match Type'] != 'duplicate match') \
	| (df['Country Code'] != 'US') & (df['Country Code'] != 'CA')]

# Create 'Duplicates' dataframe of rows with duplicate input or duplicate match
duplicates_df = df[((df['Match Type'] == 'duplicate input') | (df['Match Type'] == 'duplicate match'))]

# Optional extra dataframe
#filtered_df = df[((df['Industry'] == 'Education') | (df['Industry'] == 'Government'))]

#vid_df = df[(df['duplicate'] == True)]


# Drop unwanted columns in each dataframe
reachable_df = reachable_df.drop(columns = ['active', 'Match Type', 'Match Result', 'Source State', 'Country Name'])
not_reachable_df = not_reachable_df.drop(columns = ['active', 'Match Type', 'Match Result', 'Source State', 'Country Name'])
no_match_df = no_match_df.drop(columns = ['active', 'Match Type', 'Match Result', 'Source State', 'Country Name'])
duplicates_df = duplicates_df.drop(columns = ['active', 'Match Type', 'Match Result', 'Source State', 'Country Name'])
#filtered_df = filtered_df.drop(columns = ['active', 'Match Type', 'Match Result', 'Source State', 'Country Name', 'duplicate'])
#vid_df = vid_df.drop(columns = ['active', 'Match Type', 'Match Result', 'Source State', 'Country Name', 'duplicate'])

#----------------------------------------------------------------------------------------------------------------------

### WRITE XLSX

# Excel filename
f2 = 'Veritas_AMS_Ent_East_account_recomendations_726.xlsx'

writer = pd.ExcelWriter(f2, engine = 'xlsxwriter')
reachable_df.to_excel(writer, sheet_name = 'Reachable', index = False)
not_reachable_df.to_excel(writer, sheet_name = 'Not Reachable', index = False)
no_match_df.to_excel(writer, sheet_name = 'No Match', index = False)
duplicates_df.to_excel(writer, sheet_name = 'Duplicates', index = False)
#filtered_df.to_excel(writer, sheet_name = 'Education and Government', index = False)
#vid_df.to_excel(writer, sheet_name = 'Video Campaign', index = False)
writer.save()
