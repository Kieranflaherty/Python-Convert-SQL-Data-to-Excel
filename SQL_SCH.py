import pyodbc
import pandas as pd
import sys
from datetime import datetime

today = datetime.now()

#Establish Connection to SQL Database
server = 'EUIRLBFTTM'
database = 'BectonDickinson'
username = 'remote'
password = 'Becton2018'
connection = pyodbc.connect('DRIVER={SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password)

#Setup SQL Queries
day_l1 = "SELECT * FROM [DatabaseName].[dbo] where "blah blah blah" >= DATEADD(DAY, -1, GETDATE())
day_l2 = "SELECT * FROM [DatabaseName].[dbo] where "blah blah blah" >= DATEADD(DAY, -1, GETDATE())
day_l3 = "SELECT * FROM [DatabaseName].[dbo] where "blah blah blah" >= DATEADD(DAY, -1, GETDATE())
day_l4 = "SELECT * FROM [DatabaseName].[dbo] where "blah blah blah" >= DATEADD(DAY, -1, GETDATE())
day_l5 = "SELECT * FROM [DatabaseName].[dbo] where "blah blah blah" >= DATEADD(DAY, -1, GETDATE())
day_l6 = "SELECT * FROM [DatabaseName].[dbo] where "blah blah blah" >= DATEADD(DAY, -1, GETDATE())
day_l7 = "SELECT * FROM [DatabaseName].[dbo] where "blah blah blah" >= DATEADD(DAY, -1, GETDATE()
day_l8 = "SELECT * FROM [DatabaseName].[dbo] where "blah blah blah" >= DATEADD(MONTH, -1, GETDATE())

cursor = connection.cursor()
try:
    data = pd.read_sql(day_l1, connection)
    data.to_excel('Line_Daily_Data_'+ today.strftime('%Y_%m_%d')+'.xlsx',sheet_name='Line 1')
    print("Query Successful")
except Error as err:
    print(f"Error: '{err}'")

cursor = connection.cursor()
try:
    data2 = pd.read_sql(day_l2, connection)
    data2.to_csv('Line_Daily_Data_'+ today.strftime('%Y_%m_%d')+'.csv')
    print("Query Successful")
except Error as err:
    print(f"Error: '{err}'")

cursor = connection.cursor()
try:
    data3 = pd.read_sql(day_l3, connection)
    data.to_csv('L3_Daily_Data_'+ today.strftime('%Y_%m_%d')+'.csv')
    print("Query Successful")
except Error as err:
    print(f"Error: '{err}'")

cursor = connection.cursor()
try:
    data4 = pd.read_sql(day_l4, connection)
    data.to_csv('L4_Daily_Data_'+ today.strftime('%Y_%m_%d')+'.csv')
    print("Query Successful")
except Error as err:
    print(f"Error: '{err}'")

cursor = connection.cursor()
try:
    data5 = pd.read_sql(day_l5, connection)
    data.to_csv('L5_Daily_Data_'+ today.strftime('%Y_%m_%d')+'.csv')
    print("Query Successful")
except Error as err:
    print(f"Error: '{err}'")

cursor = connection.cursor()
try:
    data6 = pd.read_sql(day_l6, connection)
    data.to_csv('L6_Daily_Data_'+ today.strftime('%Y_%m_%d')+'.csv')
    print("Query Successful")
except Error as err:
    print(f"Error: '{err}'")

cursor = connection.cursor()
try:
    data7 = pd.read_sql(day_l7, connection)
    data.to_csv('L7_Daily_Data_'+ today.strftime('%Y_%m_%d')+'.csv')
    print("Query Successful")
except Error as err:
    print(f"Error: '{err}'")

cursor = connection.cursor()
try:
    data8 = pd.read_sql(day_l8, connection)
    data.to_csv('L8_Daily_Data_'+ today.strftime('%Y_%m_%d')+'.csv')
    print("Query Successful")
except Error as err:
    print(f"Error: '{err}'")


writer = pd.ExcelWriter('Line_Data_'+ today.strftime('%Y_%m_%d')+'.xlsx', engine = 'xlsxwriter')
data.to_excel(writer, sheet_name = 'Line 1')
data2.to_excel(writer, sheet_name = 'Line 2')
data3.to_excel(writer, sheet_name = 'Line 3')
data4.to_excel(writer, sheet_name = 'Line 4')
data5.to_excel(writer, sheet_name = 'Line 5')
data6.to_excel(writer, sheet_name = 'Line 6')
data7.to_excel(writer, sheet_name = 'Line 7')
data8.to_excel(writer, sheet_name = 'Line 8')
writer.save()
writer.close()
