### READ 'NO EXPANSION' MATCH FILE AND WRITE EXACT MATCHES TO EXCEL FILE
### IGNORES SOUNDEX F AND METAPHONE D MATCHES.

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
f = "Corp Podded West Accounts_2018-06-20_125537_output.csv"

# Create dataframe for entire match file to obtain number of columns(length).
matchfile_df = pd.read_csv(f, skiprows=None)
length = len(matchfile_df.columns)

# Select column ranges to include:
# Source Name, Match Result, Match Type, Best SID, Account, Country Code, MEI
range1 = [0,1,2,4,7,8,9,10,19,20,28]

# Range for all custom attribute columns
range2 = range(35,length)

# List of columns to include
cols = range1 + range2

# Create dataframe with selected columns
df = pd.read_csv(f, skiprows=None, usecols=cols)


#-----------------------------------------------------------------------------------------------------------------------

### FILTER DATA

# Remove rows with 'active' == 'False' from reachable and not reachable

# Create 'reachable' dataframe for 'MEI' >= 25 and 'active' == True
reachable_df = df[(df['MEI'] >= 20) & (df['active'] == True) & (df['Match Type'] != 'soundex suggestion') & (df['Match Type'] != 'left join metaphone name')]

# Create 'Not_Reachable' dataframe for 'MEI' < 25 and 'active' == True
not_reachable_df = df[(df['MEI'] < 20) & (df['active'] == True) & (df['Match Type'] != 'soundex suggestion') & (df['Match Type'] != 'left join metaphone name')]

# Create 'No_Match' dataframe for rows that require manual scrubbing
# These rows also have 'active' == False and do not contain any duplicates
no_match_df = df[(df['Match Result'] == 'no match') & (df['Match Type'] != 'duplicate input') & (df['Match Type'] != 'duplicate match') & (df['Match Type'] != 'soundex suggestion') & (df['Match Type'] != 'left join metaphone name')]

# Create 'Duplicates' dataframe of rows with duplicate input or duplicate match
duplicates_df = df[((df['Match Type'] == 'duplicate input') | (df['Match Type'] == 'duplicate match')) & (df['Match Type'] != 'soundex suggestion') & (df['Match Type'] != 'left join metaphone name')]


# Drop unwanted columns in each dataframe
reachable_df = reachable_df.drop(columns = ['active', 'Match Type', 'Match Result', 'Source State', 'Country Name'])
not_reachable_df = not_reachable_df.drop(columns = ['active', 'Match Type', 'Match Result', 'Source State', 'Country Name'])
no_match_df = no_match_df.drop(columns = ['active', 'Match Type', 'Match Result', 'Source State', 'Country Name'])
duplicates_df = duplicates_df.drop(columns = ['active', 'Match Type', 'Match Result', 'Source State', 'Country Name'])

#----------------------------------------------------------------------------------------------------------------------

### WRITE XLSX

# Excel filename
f2 = 'Corp_Podded_West_Accounts_account_recomendations_620.xlsx'

writer = pd.ExcelWriter(f2, engine = 'xlsxwriter')
reachable_df.to_excel(writer, sheet_name = 'Reachable', index = False)
not_reachable_df.to_excel(writer, sheet_name = 'Not Reachable', index = False)
no_match_df.to_excel(writer, sheet_name = 'No Match', index = False)
duplicates_df.to_excel(writer, sheet_name = 'Duplicates', index = False)
writer.save()
