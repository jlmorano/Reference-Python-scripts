################################################################
# Purpose: For each domain (particularly member institutions),
# A. report rank and total downloads for each domain by year and
# B. calculate the average rank over a 3-year period to determine Tier # for this year's and next year's cost.
# C. Format output into a file to send to institution.

## This python code accomplishes this in the following steps:
# For each institutional domain
#    1. Get data for all years for a domain, in descending order
#    2. Report average rank for current (= 2020) calendar year (average of 2016, 2017, 2018)
#    3. Report average rank for current (= 2021) calendar year (average of 2017, 2018, 2019)
#    4. Calculate the Tier and Cost based on average rank
#    5. Save report of all data to Excel

################################################################
## Need these files:
# File of rank of each domain by year 'Institutional Usage Stats-Rank_2009-2019.xlsx'
# File of domains of member-supporters 'members.xlsx'

# ***************************
## Verify these are Installed:
# latest python 3.x.x
# Anaconda, which includes pandas and numpy

################################################################
##### Prepare data
# Import pandas
import pandas as pd

# Create a function to evaluate the tier (step #4)
def tier(x):
    if x > 1 and x < 25:
        return 1
    elif x > 26 and x < 50:
        return 2
    elif x > 51 and x < 100:
        return 3
    elif x > 101 and x < 150:
        return 4
    elif x > 151 and x < 200:
        return 5
    else:
        return 6

# Create function to evaluate the cost (step 4)
def cost(x):
    if x == 1:
        return 4400
    elif x == 2:
        return 3800
    elif x == 3:
        return 3200
    elif x == 4:
        return 2500
    elif x == 5:
        return 1800
    else:
        return 1000


# Load spreadsheet of institutional stats
data = pd.read_excel('/Users/jlm394/PycharmProjects/usagestats/Institutional Usage Stats-Rank_2009-2019.xlsx')
data.head()
# Review dimensions
data.shape
# (40175, 16) #As of 12/19/2019 with 2019 usage

# Load spreadsheet of the domains of member-supporters
members = pd.read_excel('/Users/jlm394/PycharmProjects/usagestats/members.xlsx')
members.head()
# Review dimensions
members.shape
# (234, 2)

# Check years included in data
data.year.unique()
# array([2019, 2009, 2010, 2011, 2012, 2013, 2014, 2015, 2016, 2017, 2018])


##### Run through members list of domains and process for each domain

for i in members.index:
    member_domain = (members['Domain'][i])
# 1. Get data for all years for a domain, in descending order
    member_report = pd.DataFrame(data[data['domain'] == member_domain].sort_values(['year'], ascending=[False]))
    # set index to year
    member_report.set_index('year')
    #
# 2. Report average rank for current (= 2020) calendar year (average of 2016, 2017, 2018)
    currentyr = member_report.query('year == 2016 | year == 2017 | year == 2018')
    currentyr_rank = currentyr["rank"].mean()
    #
# 3. Report average rank for current (= 2021) calendar year (average of 2017, 2018, 2019)
    nextyr = member_report.query('year == 2017 | year == 2018 | year == 2019')
    nextyr_rank = nextyr["rank"].mean()
    # Gather all of the data for averaged rank and cost into a dataframe
    list = {'CY':['2020', '2021'],
            '3-Yr_Rank_Average':[currentyr_rank, nextyr_rank]}
    # Create the pandas DataFrame
    CYcost = pd.DataFrame(list)
    # set index to CY
    CYcost.set_index('CY')
    #
# 4. Calculate the Tier and Cost based on average rank
    # Calculate the tier
    CYcost['Tier'] = CYcost['3-Yr_Rank_Average'].apply(tier)
    # Calculate the cost
    CYcost['Cost'] = CYcost['Tier'].apply(cost)
    #
# 5. Save report of all data to Excel
    # Write dataframes to same sheet in excel
    writer = pd.ExcelWriter(member_domain + '_CY2020.xlsx', engine ='xlsxwriter')
    member_report.to_excel(writer, sheet_name ='Sheet1', index = False)
    CYcost.to_excel(writer, sheet_name ='Sheet1', startrow = 15, index = False)
    writer.save()

