### Reading exported CSV and selecting only required columns into a XLS file
# Import modules 
import pandas as pd
import os

# Establish original csv to read from: "base"
base = pd.read_csv('Athena RFP CHS Audience_2018-09-06_130621_output.csv')

# Establish range of columns needed - will need to keep Match Type and Active status for filters 
length = len(base.columns)
pos = [0, 1, 2, 8, 9, 10, 19, 21, 28]
# From column 35 on is all custom attributes, so must include all columns that appear
pos2 = list(range(35, length))
colnames = pos + pos2

final = pd.read_csv('Athena RFP CHS Audience_2018-09-06_130621_output.csv', usecols=colnames)

# Filter and then create new dataframe to create new tab in xls 
reachable_filter = (final['active']==True) & (final['MEI']>=20)
reachable_df = final[reachable_filter].sort_values("MEI", inplace = False, ascending=False)

not_reachable_filter = (final['MEI']<20) & (final['active']==True)
not_reachable_df = final[not_reachable_filter].sort_values('MEI', inplace=False, ascending=False)

# No matches can also be duplicates so need to figure out a way that eliminates the duplicate rows from no match tab
no_match_filter = (final['active']==False) & (final['Match Type']!='duplicate match') & (final['Match Type']!='duplicate input')
no_match_df = final[no_match_filter].sort_values('Source Name', inplace=False, ascending=False)

# Created this to place duplicate results in duplicates tab
duplicate_filter = (final['Match Type']=='duplicate match') | (final['Match Type']=='duplicate input')
duplicate_df = final[duplicate_filter].sort_values('Source Name', inplace=False, ascending=False) 

# Set new file path/directory for newly created xls files
os.chdir("/Users/jshek/Desktop/completedxls")

# Drops columns after satisfying conditions
dropcols = ['active']
reachable_df.drop(dropcols, axis=1, inplace=True)
not_reachable_df.drop(dropcols, axis=1, inplace=True)
no_match_df.drop(dropcols, axis=1, inplace=True)
duplicate_df.drop(dropcols, axis=1, inplace=True)

from pandas import ExcelWriter
writer = pd.ExcelWriter('Athena RFP CHS Audience_2018-09-06_130621_output.xlsx', engine='xlsxwriter')

reachable_df.to_excel(writer, sheet_name='Reachable', index=False)
not_reachable_df.to_excel(writer, sheet_name='Not Reachable', index=False)
no_match_df.to_excel(writer, sheet_name='No Match', index=False)
duplicate_df.to_excel(writer, sheet_name='Duplicates', index=False)

writer.save()
