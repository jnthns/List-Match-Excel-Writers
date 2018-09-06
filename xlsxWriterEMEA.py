### READ EXPORTED MATCH FILE AND WRITE ONLY ROWS WITH EMEA COUNTRIES TO
### REACHABLE, NOT REACHABLE, NO MATCH TABS.  WRITE NON EMEA ROWS TO DUPLICATES TAB

import pandas as pd
import sys  

# Changed default string encoding from 'ascii' to 'utf8'
reload(sys)  
sys.setdefaultencoding('utf8')

### READ CSV 

# Match filename
f = "Lithium EMEA Ads_2018-08-29_130178_output (1).csv"

# Create dataframe for entire match file to obtain number of columns(length).
matchfile_df = pd.read_csv(f, skiprows=None)
length = len(matchfile_df.columns)

# Select column ranges to include:
# Source Name, Match Result, Match Type, Best SID, Account, Country Code, MEI
range1 = [0,1,2,4,7,8,9,10,14,19,20,21,28, 31]

# Range for all custom attribute columns
range2 = range(35,length)

# List of columns to include
cols = range1 + range2

# Create dataframe with selected columns
df = pd.read_csv(f, skiprows=None, usecols=cols)

EMEA_Countries = ['AL', 'DZ', 'AD', 'AO', 'AT', 'BH', 'BY', 'BE', 'BJ', 'BA', 'BW', \
'BG', 'BF', 'BI', 'CM', 'CV', 'CF', 'TD', 'KM', 'HR', 'CY', 'CZ', 'CD', 'DK', 'DJ', \
'EG', 'GQ', 'ER', 'EE', 'ET', 'FO', 'FI', 'FR', 'GA', 'GM', 'GE', 'DE', 'GH', 'GI', \
'GR', 'GG', 'GN', 'GW', 'HU', 'IS', 'IR', 'IQ', 'IE', 'IM', 'IL', 'IT', 'CI', 'JE', \
'JO', 'KE', 'KW', 'LV', 'LB', 'LS', 'LR', 'LY', 'LI', 'LT', 'LU', 'MK', 'MG', 'MW', \
'ML', 'MT', 'MR', 'MU', 'MD', 'MC', 'ME', 'MA', 'MZ', 'NA', 'NL', 'NE', 'NG', 'NO', \
'OM', 'PS', 'PL', 'PT', 'QA', 'RO', 'RW', 'SM', 'ST', 'SA', 'SN', 'RS', 'SK', 'SI', \
'SO', 'ZA', 'ES', 'SD', 'SZ', 'SE', 'CH', 'SY', 'TZ', 'TG', 'TN', 'TR', 'UG', 'UA', \
'AE', 'GB', 'VA', 'EH', 'YE', 'ZM', 'ZW']

#-----------------------------------------------------------------------------------------------------------------------

def EmeaCountry(country):

	if country in EMEA_Countries:

		return True

	else:

		return False

#-----------------------------------------------------------------------------------------------------------------------

def EmeaCountryColumn(country):

	df['EMEA'] = country.apply(EmeaCountry)

	return df 

#-----------------------------------------------------------------------------------------------------------------------

# Create EMEA Country column with value of True or False

df = EmeaCountryColumn(df['Country Code'])

# FILTER DATA

# Remove rows with 'active' == 'False' and 'EMEA' == 'False' from reachable and not reachable

# Create 'reachable' dataframe for 'MEI' >= 25 and 'active' == True
reachable_df = df[(df['MEI'] >= 20) & (df['active'] == True) & (df['EMEA'] == True)]

# Create 'Not_Reachable' dataframe for 'MEI' < 25 and 'active' == True
not_reachable_df = df[(df['MEI'] < 20) & (df['active'] == True) & (df['EMEA'] == True)]

# Create 'No_Match' dataframe for rows that require manual scrubbing
# These rows also have 'active' == False and do not contain any duplicates
no_match_df = df[(df['Match Result'] == 'no match') & (df['Match Type'] != 'duplicate input') \
	& (df['Match Type'] != 'duplicate match') & (df['EMEA'] == True)]

# Create 'Duplicates' dataframe of rows with duplicate input or duplicate match
duplicates_df = df[((df['Match Type'] == 'duplicate input') | (df['Match Type'] == 'duplicate match') \
	| (df['EMEA'] == False))]


# Drop unwanted columns in each dataframe
reachable_df = reachable_df.drop(columns = ['active', 'Match Type', 'Match Result', 'Source State', 'Country Name', 'EMEA'])
not_reachable_df = not_reachable_df.drop(columns = ['active', 'Match Type', 'Match Result', 'Source State', 'Country Name', 'EMEA'])
no_match_df = no_match_df.drop(columns = ['active', 'Match Type', 'Match Result', 'Source State', 'Country Name', 'EMEA'])
duplicates_df = duplicates_df.drop(columns = ['active', 'Match Type', 'Match Result', 'Source State', 'Country Name', 'EMEA'])

#----------------------------------------------------------------------------------------------------------------------

### WRITE XLSX

# Excel filename
f2 = 'Lithium_EMEA_Ads_account_recomendations_829.xlsx'

writer = pd.ExcelWriter(f2, engine = 'xlsxwriter')
reachable_df.to_excel(writer, sheet_name = 'Reachable', index = False)
not_reachable_df.to_excel(writer, sheet_name = 'Not Reachable', index = False)
no_match_df.to_excel(writer, sheet_name = 'No Match', index = False)
duplicates_df.to_excel(writer, sheet_name = 'Duplicates', index = False)
writer.save()

