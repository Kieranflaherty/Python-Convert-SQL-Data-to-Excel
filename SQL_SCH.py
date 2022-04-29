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
day_l1 = "SELECT * FROM [BectonDickinson].[dbo].[ProcessDataOutput0102] where S10232_DataTimestamp >= DATEADD(DAY, -1, GETDATE()) ORDER BY S10232_DataTimestamp DESC;"
day_l2 = "SELECT * FROM [BectonDickinson].[dbo].[ProcessDataOutput0202] where S20212_DataTimestamp >= DATEADD(DAY, -1, GETDATE()) ORDER BY S20212_DataTimestamp DESC;"
day_l3 = "SELECT * FROM [BectonDickinson].[dbo].[ProcessDataOutput0302] where S30232_DataTimestamp >= DATEADD(DAY, -1, GETDATE()) ORDER BY S30232_DataTimestamp DESC;"
day_l4 = "SELECT * FROM [BectonDickinson].[dbo].[ProcessDataOutput0401] where S40132_DataTimestamp >= DATEADD(DAY, -1, GETDATE()) ORDER BY S40132_DataTimestamp DESC;"
day_l5 = "SELECT * FROM [BectonDickinson].[dbo].[ProcessDataOutput0501] where S50132_DataTimestamp >= DATEADD(DAY, -1, GETDATE()) ORDER BY S50132_DataTimestamp DESC;"
day_l6 = "SELECT * FROM [BectonDickinson].[dbo].[ProcessDataOutput0602] where S60232_DataTimestamp >= DATEADD(DAY, -1, GETDATE()) ORDER BY S60232_DataTimestamp DESC;"
day_l7 = "SELECT * FROM [BectonDickinson].[dbo].[ProcessDataOutput0702] where S70232_DataTimestamp >= DATEADD(DAY, -1, GETDATE()) ORDER BY S70232_DataTimestamp DESC;"
day_l8 = "SELECT * FROM [BectonDickinson].[dbo].[ProcessDataOutput0803] where L8_DataTimestamp >= DATEADD(MONTH, -1, GETDATE()) ORDER BY L8_DataTimestamp DESC;"

cursor = connection.cursor()
try:
    data = pd.read_sql(day_l1, connection)
#    data.to_excel('Line_Daily_Data_'+ today.strftime('%Y_%m_%d')+'.xlsx',sheet_name='Line 1')
    print("Query Successful")
except Error as err:
    print(f"Error: '{err}'")

cursor = connection.cursor()
try:
    data2 = pd.read_sql(day_l2, connection)
#    data2.to_csv('Line_Daily_Data_'+ today.strftime('%Y_%m_%d')+'.csv')
    print("Query Successful")
except Error as err:
    print(f"Error: '{err}'")

cursor = connection.cursor()
try:
    data3 = pd.read_sql(day_l3, connection)
#    data.to_csv('L3_Daily_Data_'+ today.strftime('%Y_%m_%d')+'.csv')
    print("Query Successful")
except Error as err:
    print(f"Error: '{err}'")

cursor = connection.cursor()
try:
    data4 = pd.read_sql(day_l4, connection)
#    data.to_csv('L4_Daily_Data_'+ today.strftime('%Y_%m_%d')+'.csv')
    print("Query Successful")
except Error as err:
    print(f"Error: '{err}'")

cursor = connection.cursor()
try:
    data5 = pd.read_sql(day_l5, connection)
#    data.to_csv('L5_Daily_Data_'+ today.strftime('%Y_%m_%d')+'.csv')
    print("Query Successful")
except Error as err:
    print(f"Error: '{err}'")

cursor = connection.cursor()
try:
    data6 = pd.read_sql(day_l6, connection)
#    data.to_csv('L6_Daily_Data_'+ today.strftime('%Y_%m_%d')+'.csv')
    print("Query Successful")
except Error as err:
    print(f"Error: '{err}'")

cursor = connection.cursor()
try:
    data7 = pd.read_sql(day_l7, connection)
#    data.to_csv('L7_Daily_Data_'+ today.strftime('%Y_%m_%d')+'.csv')
    print("Query Successful")
except Error as err:
    print(f"Error: '{err}'")

cursor = connection.cursor()
try:
    data8 = pd.read_sql(day_l8, connection)
#    data.to_csv('L8_Daily_Data_'+ today.strftime('%Y_%m_%d')+'.csv')
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
