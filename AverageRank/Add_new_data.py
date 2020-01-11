################################################################
# Purpose: Add to the list of Institutional Usage Stats-Rank since 2009
# the new year's rank, total, and total by category for each domain

## This python code accomplishes this in the following steps:
#    1. Lookup by domain,
#       a. The rank and total downloads in one spreadsheet
#       b. The total downloads by category in another spreadsheet
#    2. Write the data to the database of all downloads and rank by domain

################################################################
## Need these files:
# 'Final Rank by 2019 downloads_top_1000_int'
# 'Institutional Usage Stats-Rank_2009-2018.xlsx'

# ***************************
## Verify these are Installed:
# latest python 3.x.x
# Anaconda, which includes pandas and numpy

################################################################
##### Prepare data
# Import pandas
import pandas as pd

# Check current directory
# os.getcwd()
#'/Users/jlm394/PycharmProjects/usagestats'
# os.chdir('/Users/jlm394/PycharmProjects/usagestats')

# Load worksheet of institutional stats since 2009
data = pd.read_excel('/Users/jlm394/PycharmProjects/usagestats/Institutional Usage Stats-Rank_2009-2018.xlsx')

# Load worksheet of Final Rank by domain for 2019
rank = pd.read_excel('Final Rank by 2019 downloads_top_1000_inst.xlsx', sheet_name="rank_python")
# Load worksheet of Final Rank by domain for 2019
category = pd.read_excel('Final Rank by 2019 downloads_top_1000_inst.xlsx', sheet_name="category_python")

#Combine rank and category by the domain
results=rank.merge(category, on="domain")

# Rename columns
results.rename(columns = {'downloads':'total', 'RankbyIP':'RankByIP'}, inplace = True)

# Reorder columns to match 'data'
# Stopped working for unknown reason
#cols = list(results.columns.values) #Get all column names and put in a list
#cols.pop(cols.index('RankByIP')) # Remove RankByIP from ordered list
#results = results[cols+['RankByIP']] # Move RankByIP to end of list and columns

results.columns = ['year', 'rank', 'domain', 'total',
              'Astrophysics', 'Cond_Matter_Physics', 'Computer_Science', 'Economics',
       'Electrical_Engineering_and_Systems_Science',
       'High_Energy_Physics', 'Mathematics', 'Other_Physics',
       'Quantitative_Biology', 'Quantitative_Finance', 'Statistics',
              'RankByIP']

# Append 'results' to 'data'
data.shape
# (39177, 16)
new = results.append(data, ignore_index = True)
new.shape
# (40175, 16)

# Write 'new' to excel
export_excel = new.to_excel (r'/Users/jlm394/PycharmProjects/usagestats/Institutional Usage Stats-Rank_2009-2019.xlsx', index = None, header=True)
