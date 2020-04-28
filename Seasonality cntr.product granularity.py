# -*- coding: utf-8 -*-
"""Classical time series decomposition to get trend, seasonality and remainder. Assumes
multiplicative decomposition, that is, y_t = T_t * S_t * R_t

Created on Tue Apr 28 16:14:31 2020

@author: Artyom, Simene
"""
# =============================================================================
# Preamble
# =============================================================================
import pandas as pd
import numpy as np
import sys
import matplotlib.pyplot as plt
import datetime
import os
from calendar import monthrange

# =============================================================================
#  User-defined functions
# =============================================================================

# =============================================================================
# Predefined variables
# =============================================================================

path = r'C:\Users\.....'
input_filename = 'Demand seasonality project Python.xlsx'
output_filename = 'Seasonality_ratios.xlsx'

# =============================================================================
# Initial cleaning of data
# =============================================================================
# Open input_filename
df_original = pd.read_excel(path+input_filename, sheet_name = 'Sheet1', header = 0, usecols = 'F:J')

# Take copy of original frame
df = df_original.copy()

# Set 'NODATA' to zero, and drop all  and negative values
df.loc[df['Value'] == 'NODATA','Value'] = 0
df['Value'] = df['Value'].astype(float)
df = df.loc[df['Value'] > 0]

# Drop nan
df = df.loc[df['Value'].isna() != True]

# Ensure correct sorting
df.sort_values(by=['Country','Product','Year','Month'], ascending=[True, True, True, True], inplace=True)
df.reset_index(inplace=True)

# Format date
df['Day'] = 1
df['date'] = pd.to_datetime(df[['Year','Month','Day']],format='%Y%B%d')
df.drop(columns = 'Day',inplace=True)

# Add missing dates. Addes missing months (resample 'MS') and allocates blank values to all columns in 
#a missing dates row (asfreq() function), then drop country and product index
df = df.set_index('date').groupby(['Country', 'Product']).resample('MS').asfreq().reset_index(level=0, drop=True)
df = df.reset_index(level=0,drop=True).reset_index()

# Format the row for missing date.
df['Year'] = df['date'].apply(lambda x: x.year)
df['Month'] = df['date'].apply(lambda x: x.month)
df['Country'] = df['Country'].ffill()
df['Product'] = df['Product'].ffill()
df['Value'] = df['Value'].interpolate(method='linear', order=1)

# =============================================================================
#  Time-series decomposition
# =============================================================================
#### Calculate the trend by doing a  12 Moving Average with a further centralization around months 6 and 7
df['12MA'] = df.groupby(['Country', 'Product'])['Value'].rolling(window=12,axis=0).mean().shift(-6).reset_index(drop=True)
df['Trend'] = df.groupby(['Country', 'Product'])['12MA'].rolling(window=2,axis=0).mean().shift(-1).reset_index(drop=True)


#### Monthly seasonality taking y_t / T_t
df['Monthly Seasonality'] =  df['Value'] / df['Trend']
df.dropna(inplace=True)

#Scale seasonality ratios to number of datapoints each year
# Group and sum monthly seasonality, rename column and merge to dataframe
sum_seasonality = df.groupby(['Country','Product','Year'])['Monthly Seasonality'].sum().reset_index()
sum_seasonality.rename(columns={'Monthly Seasonality': 'Sum Monthly Seasonality'}, inplace=True)
df = pd.merge(df, sum_seasonality, on=['Country', 'Product','Year'], how='left')

# Find number of datapoints within one year. First and last reported year may have less points than 12 months
df['datapoints per year'] = df.groupby(['Country','Product','Year'])['Month'].transform('nunique')

# Scale seasonality to add up to number of datapoints per year
df['Scaled seasonality'] = df['Monthly Seasonality'] / df['Sum Monthly Seasonality'] * df['datapoints per year']

#consider only last 10 years of seasonality in the spreadsheet.

df = df.loc[df['Year']>=datetime.date.today().year-10]


            
## Drop max and min value from each group of Country, Month and Product
average_seasonality = df.copy()

    # Here we in the "average_seasonality" intermidiary table for each country/product/month drop 1 maximum value for each country,product and month
    #then we do the same for the minimum values 
average_seasonality = average_seasonality.drop(average_seasonality.sort_values('Scaled seasonality',ascending=True).groupby(['Country','Month','Product']).tail(1).index)
average_seasonality = average_seasonality.drop(average_seasonality.sort_values('Scaled seasonality',ascending=False).groupby(['Country','Month','Product']).tail(1).index)

#weight according to years (more recent years get more weight)
    #get only unique years per every country, prdocut and month
sum_all_years=average_seasonality[['Country', 'Product','Month','Year']].drop_duplicates()
    #sum unique years
sum_all_years=sum_all_years.groupby(['Country', 'Product','Month'])['Year'].sum().reset_index() 
sum_all_years.rename(columns={'Year': 'Sumyear'}, inplace=True)
    #year/sumyear gives a weight per yer. Closer the year to the actual date- more weight it gets
average_seasonality = pd.merge(average_seasonality, sum_all_years, on=['Country', 'Product', 'Month'], how='left')
average_seasonality['Weight%']=average_seasonality['Year']/average_seasonality['Sumyear']
average_seasonality.drop(columns = 'Sumyear',inplace=True)

# Get average seasonality for a given month
average_seasonality['seasonality_scaled_weighted']=average_seasonality['Scaled seasonality']*average_seasonality['Weight%']
average_seasonality = average_seasonality.groupby(['Country','Month','Product'])['seasonality_scaled_weighted'].sum().reset_index()
average_seasonality.rename(columns={'seasonality_scaled_weighted': 'Avg Scaled weighted seasonality'}, inplace=True)

#scale the final result to 12
    # create a column with sum of all seasonalities
average_seasonality2 = average_seasonality.groupby(['Country','Product'])['Avg Scaled weighted seasonality'].sum().reset_index()
average_seasonality2.rename(columns={'Avg Scaled weighted seasonality': 'sumseasonalities'}, inplace=True)
    # merge with existing datasheet
average_seasonality = pd.merge(average_seasonality, average_seasonality2, on=['Country', 'Product'], how='left').reset_index()
    #scale to 12 by x/sum(x)*12 
average_seasonality['Avg Scaled weighted seasonality']=average_seasonality['Avg Scaled weighted seasonality']/average_seasonality['sumseasonalities']*12
average_seasonality.drop(columns= 'sumseasonalities', inplace=True)
del(average_seasonality2)
#sort
average_seasonality.sort_values(by=['Product','Country','Month'], ascending=[True, True, True], inplace=True)
average_seasonality.reset_index(inplace=True)

# # String together seasonality to time series
df = pd.merge(df, average_seasonality, how='left', on=['Country','Month','Product'])
final = df[['date', 'Year', 'Month', 'Country', 'Product', 'Value','Avg Scaled weighted seasonality']]


# =============================================================================
#  Running some final checks and writing to a new file on the server
# =============================================================================

#creating a check dataframe for KSA, making sure all fuels are present
dfksa=final[final['Country']=='Saudi Arabia']


#wrtie into destination excel
writer=pd.ExcelWriter(path+output_filename, engine='xlsxwriter')
final.to_excel(writer, sheet_name='Sheet1')
writer.save()
