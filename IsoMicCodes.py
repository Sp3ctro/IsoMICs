# -*- coding: utf-8 -*-
"""
This script downloads the latest excel file from the ISO Group which lists MIC code registrations

"""
import time
import datetime
from dateutil.relativedelta import relativedelta
import pandas as pd

# read in a CSV file from the ISO Group's Website
c= pd.read_csv('https://www.iso20022.org/sites/default/files/ISO10383_MIC/ISO10383_MIC.csv', encoding = 'iso-8859-1')

# replace null values with "nan" to avoid issues with converting to a dataframe
c = c.where(pd.notnull(c), "nan")

# pull the CREATION DATE column values into a list wihch can be manipulated
oldDates = c['CREATION DATE'].tolist()

# for each value in oldDates...
for index, value in enumerate(oldDates):
    try :
        # overwrite it with the date in datetime object form
        oldDates[index] = time.strptime(oldDates[index], '%B %Y')
    except :
        # handle the single value in the file which breaks the regex
        oldDates[index] = time.strptime('JUNE 2005', '%B %Y')

# replace the dataframe column with the datetime object values and sort by date
c['CREATION DATE'] = oldDates
c = c.sort_values(['CREATION DATE'])

# get today's date, subtract 4 months from it
TODAY = datetime.date.today()
lastupdate_time = time.strptime(str(TODAY - relativedelta(months=4)), '%Y-%m-%d')

# make a new dataframe where with new MIC codes created in the last 4 months
d = c.where(c['CREATION DATE'] > lastupdate_time)
d = d.dropna()

# write both dataframs to an excel file in separate sheets
writer = pd.ExcelWriter('PythonOutput.xlsx')
c.to_excel(writer, 'Latest File')
d.to_excel(writer, 'Added in the past 4 months')
writer.save()