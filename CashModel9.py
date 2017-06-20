#Import Libraries
import pandas as pd
import numpy as np
from datetime import datetime,date,timedelta
import types
import string

# default='warn' *Removes warning from timedelta merger
pd.options.mode.chained_assignment = None

#Dataset location
Location = r'D:\Users\schades\python\Cash Model\BC_Data.xlsx'

#Import billing and colection data from workbook
bd = pd.read_excel(Location,sheetname='Billing')
cd = pd.read_excel(Location,sheetname='Collections')
mp = pd.read_excel(Location,sheetname='Map')
wm = pd.read_excel(Location,sheetname='WeekMap')
ol = pd.read_excel(Location,sheetname='Sales OL')
bill_ol = pd.read_excel(Location,sheetname='OL Bill')

#Select certain columns
bdf = pd.DataFrame(bd, columns=['Project ID', 'Invoice Number','Invoice Date','Invoice Amount'])
cdf = pd.DataFrame(cd, columns=['Receipt Date', 'Invoice ID', 'Transaction Crncy Amount'])

#Week Filter
Week_Filter = int(input('Enter Current Week:'))

#Add ProjectID Level 1 to DataFrame
bdf['ProjectLvl1'] = bd['Project ID'].apply(lambda x: str(x)[:6])
#Convert Project ID in mp to object
mp['Project ID'] = mp['Project ID'].apply(lambda x: str(x))
#Convert Date in wm to Datetime
wm['Date'] = pd.to_datetime(wm['Date'])

#Test Data_Start ************

#Merge cdf and bdf into one DataFrame
Test_Data=bdf.merge(cdf, how='left', left_on='Invoice Number', right_on='Invoice ID')
#Find Average days it takes to pay invoice(Had to use '-' because values came out negative)
Test_Data['DaysToPay']= -(Test_Data['Invoice Date']-Test_Data['Receipt Date'])
#Takes DaysToPay values that are not null(This forms Test data)
Test_Data1 = Test_Data[pd.notnull(Test_Data['DaysToPay'])]
#Convert TimeDelta from DaysToPay to numeric value
Test_Data1['DaysToPay'] = pd.to_timedelta(Test_Data1['DaysToPay']) / pd.Timedelta(days=1)
#Groupby mean
result = Test_Data1.groupby('ProjectLvl1')[['DaysToPay']].mean()
#Remove Float from mean
result['DaysToPay'] = result['DaysToPay'].astype(np.int64)

#Test Data_End ****************

#OUTLOOK DATA_Start***********
#Well this is completely messed up, are you sure we should do this?\

a = bd.TD.reset_index()

#groupby Month
ol_mo = ol.set_index('Project ID').unstack()
#remove index
ol_ri = ol_mo.reset_index()
#Add in Billing Date Column by project
ol_mo1 = pd.merge(ol_ri,bill_ol,left_on='level_0',right_on='Sales Month')
#Change name of 0 column to OL Sales
ol_mo1.columns.values[2] = 'Sales'
#delete all zero values in OL Sales
ol_mo1 = ol_mo1.drop(ol_mo1[ol_mo1.Sales < 1].index)
#convert Bill Date to datetime
ol_mo1['Bill Date'] = pd.to_datetime(ol_mo1['Bill Date'])
#remove time from date
ol_mo1['Bill Date'] = ol_mo1['Bill Date'].apply(lambda x: x.date())
#ProjectID to object
ol_mo1['Project ID'] = ol_mo1['Project ID'].apply(lambda x: str(x))
#merge days to pay into sheet
ol_mo1 = pd.merge(ol_mo1,result,left_on='Project ID',right_index=True)
#Convert bill data to datetime
ol_mo1['Bill Date'] = pd.to_datetime(ol_mo1['Bill Date'])
#add est pay date, week and project Name into OL data
ol_mo1['Estimated_Pay_Date'] = ol_mo1['Bill Date'] + pd.to_timedelta(ol_mo1['DaysToPay'],unit='D')
ol_mo1 = pd.merge(ol_mo1,wm,left_on='Estimated_Pay_Date',right_on='Date',right_index=False)
ol_mo1 = pd.merge(ol_mo1,mp,left_on='Project ID',right_on='Project ID',right_index=False)
#Create Description Column
ol_mo1['Invoice Number'] = ol_mo1['Sales Month'] + ' Sales'
#Clean Columns in ol_mo1/Change Names
ol_mo1 = ol_mo1[['Project ID','Name','Invoice Number','Bill Date','Sales','Estimated_Pay_Date','Week','Month']]
ol_mo1.columns.values[3] = 'Invoice Date'
ol_mo1.columns.values[4] = 'Invoice Amount'
#Remove Time from Date
ol_mo1['Estimated_Pay_Date'] = ol_mo1['Estimated_Pay_Date'].apply(lambda x: x.date())
ol_mo1['Invoice Date'] = ol_mo1['Invoice Date'].apply(lambda x: x.date())
#Filter OL based on Criteria
ol_mo1 = ol_mo1.drop(ol_mo1[ol_mo1.Week < Week_Filter + 2].index)

#OUTLOOK DATA_End***********

#Bill Data _Start *******

#List of Invoices that are don't have matching pay date
Bill_Data1 = Test_Data[Test_Data['DaysToPay'].isnull()][['ProjectLvl1','Invoice Number','Invoice Date','Invoice Amount']]
#convert Invoice Date to datetime
Bill_Data1['Invoice Date'] = pd.to_datetime(Bill_Data1['Invoice Date'])
#Merge result and Bill Data on ProjectLvl1
Bill_Data1 = pd.merge(Bill_Data1,result,left_on='ProjectLvl1',right_index=True)
#Create Estimated Pay Date Column
Bill_Data1['Estimated_Pay_Date'] = Bill_Data1['Invoice Date'] + pd.to_timedelta(Bill_Data1['DaysToPay'],unit='D')
#Remove Time from Date
Bill_Data1['Estimated_Pay_Date'] = Bill_Data1['Estimated_Pay_Date'].apply(lambda x: x.date())
Bill_Data1['Invoice Date'] = Bill_Data1['Invoice Date'].apply(lambda x: x.date())
#Convert Estimated_Pay_Date to Date Time
Bill_Data1['Estimated_Pay_Date'] = pd.to_datetime(Bill_Data1['Estimated_Pay_Date'])
#Add Project Name to DataFrame
Bill_Data1 = pd.merge(Bill_Data1,mp,left_on='ProjectLvl1',right_on='Project ID',right_index=False)
#Add Week number
Bill_Data1 = pd.merge(Bill_Data1,wm,left_on='Estimated_Pay_Date',right_on='Date',right_index=False)
#Remove Time from Date(round2)
Bill_Data1['Estimated_Pay_Date'] = Bill_Data1['Estimated_Pay_Date'].apply(lambda x: x.date())
#Re-arrange and select columns
Bill_Data1 = Bill_Data1[['Project ID','Name','Invoice Number','Invoice Date','Invoice Amount','Estimated_Pay_Date','Week','Month']]
#Add in OL to Bill Data
frames = [Bill_Data1,ol_mo1]
Bill_Data2 = pd.concat(frames)
#Change Column Names
Bill_Data2.columns.values[4] = 'Invoice_Amount'

#Bump up Data within 3 weeks to current week, 2 weeks +1, 1 week +2
Bill_Data2.Week[Bill_Data2.Week <= (Week_Filter - 3)] = Week_Filter
Bill_Data2.Week[Bill_Data2.Week <= (Week_Filter - 2)] = Week_Filter + 1
Bill_Data2.Week[Bill_Data2.Week <= (Week_Filter - 1)] = Week_Filter + 2

#Filter out invoice less than zero, less than week, and outside 13-week forecast
Bill_Data2 = Bill_Data2.drop(Bill_Data2[Bill_Data2.Invoice_Amount < 1].index)
Bill_Data2 = Bill_Data2.drop(Bill_Data2[Bill_Data2.Week < Week_Filter].index)
Bill_Data2 = Bill_Data2.drop(Bill_Data2[Bill_Data2.Week > Week_Filter + 12].index)

#Write Data to Excel with multiple sheets
writer = pd.ExcelWriter(r'D:\Users\schades\python\Cash Model\BC_Data1.xlsx')
Bill_Data2.to_excel(writer,'Bill Data',index=False)
Test_Data1.to_excel(writer,'Test Data',index=False)
Bill_Data1.to_excel(writer,'Bill Data NF',index=False)
writer.save()
print('Done')
