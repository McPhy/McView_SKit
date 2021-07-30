#!/usr/bin/env python
#Learn how this works here: http://youtu.be/pxofwuWTs7c
import requests
import json
import pprint
import pandas as pd
import pygsheets
import gspread
from gspread_dataframe import set_with_dataframe

#t2mdevid = d3207bb8-aa23-4d43-a677-5982c1eb8997
#t2mtoken = 4dxbErD0xMkky13rH2DofsBaN9WYcmscfX5ihfk8EyMUC5DkYA  Token CNR EWON
#CNR lyon Ewon ID = 897989 , este valor se encuentra con el post de <Getstatus>


#t2mtoken = QqvKbj3PrW4ktcukDQXTngkH1MBnIdAOCZ7dR2XBXEQvqLJtoA  Token HRS1-SMTAG
# HRS1-SMTAG ewon ID = 749009


# --- POST TIYPES ----

# GETSTATUS POST
#Authentication = {'t2mdevid': 'd3207bb8-aa23-4d43-a677-5982c1eb8997','t2mtoken': 'QqvKbj3PrW4ktcukDQXTngkH1MBnIdAOCZ7dR2XBXEQvqLJtoA'}
#respuesta = requests.post('https://data.talk2m.com/getstatus', data=Authentication)
#print(respuesta.text)
#pprint.pprint(respuesta.json())
#print(respuesta.json().get(''))

# GETDATA POST
#Authentication = {'t2mdevid': 'd3207bb8-aa23-4d43-a677-5982c1eb8997','t2mtoken': 'QqvKbj3PrW4ktcukDQXTngkH1MBnIdAOCZ7dR2XBXEQvqLJtoA', 'ewonId':'749009'}
#respuesta = requests.post('https://data.talk2m.com/getdata', data=Authentication)
#print(respuesta.text)
#pprint.pprint(respuesta.json())


# GETEWONS POST
#Authentication = {'t2mdevid': 'd3207bb8-aa23-4d43-a677-5982c1eb8997','t2mtoken': '4dxbErD0xMkky13rH2DofsBaN9WYcmscfX5ihfk8EyMUC5DkYA'}
#respuesta = requests.post('https://data.talk2m.com/getewons', data=Authentication)
#print(respuesta.text)
#pprint.pprint(respuesta.json())



# Sync DATA POST
#Authentication = {'t2mdevid': 'd3207bb8-aa23-4d43-a677-5982c1eb8997','t2mtoken': 'QqvKbj3PrW4ktcukDQXTngkH1MBnIdAOCZ7dR2XBXEQvqLJtoA', 'createTransaction':'true', 'ewonId':'749009'}
#respuesta = requests.post('https://data.talk2m.com/syncdata', data=Authentication)
#pprint.pprint(respuesta.json())
#print(respuesta.text)
"""
TransactionID = respuesta.json().get('transactionId')
FLAGMoreData = respuesta.json().get('moreDataAvailable')
HistoryTag1 = respuesta.json().get('ewons')[0].get('tags')[0].get('history')
HistoryTag2 = respuesta.json().get('ewons')[0].get('tags')[1].get('history')
HistoryTag3 = respuesta.json().get('ewons')[0].get('tags')[2].get('history')
HistoryTag4 = respuesta.json().get('ewons')[0].get('tags')[3].get('history')
print(TransactionID)
print(FLAGMoreData)
print(HistoryTag1)
print(HistoryTag2)
print(HistoryTag3)
print(HistoryTag4)
pprint.pprint(respuesta.json())
#pprint.pprint(respuesta.json().getdata(''))
"""
# Walid Talk2M Developer ID is : 87e76678-d393-4bdc-88a4-08cf164b4944
#API Key HRS-RUNGIS : PexqYJxwfBRTG86Ae8CyxNcOeebAuQerubVuV4YFYnw9zsqjmA

#API Key HRS-ROUEN :w0DRk9x2GuYXOERMAUUUNc0PDaJRI6Xf0CkDgGbjSONvx5ilQA

# LastTransactionSTR = str(49425)
# # Sync DATA POST used in Mcview Server.
# Authentication = {'t2mdevid': 'd3207bb8-aa23-4d43-a677-5982c1eb8997','t2mtoken': 'QqvKbj3PrW4ktcukDQXTngkH1MBnIdAOCZ7dR2XBXEQvqLJtoA', 'createTransaction':'true','lastTransactionId':LastTransactionSTR}
# respuesta = requests.post('https://data.talk2m.com/syncdata', data=Authentication)
# respuestajson = respuesta.json()
# print(respuesta)

#########################################################################################################################################################
exchange_table = pd.read_excel('McView Exchange Table_ HRSRouen.xlsx', sheet_name ='in', header=0)
Authentication = {'t2mdevid': '87e76678-d393-4bdc-88a4-08cf164b4944','t2mtoken': 'w0DRk9x2GuYXOERMAUUUNc0PDaJRI6Xf0CkDgGbjSONvx5ilQA', 'ewonId':'375732'}
respuest = requests.post('https://data.talk2m.com/getdata', data=Authentication)
# Syncdata ====> last transaction ID

#gc = pygsheets.authorize(service_file='mcview-starter-kit.json')
#spreadsheet = gc.open_by_key('1sHhHoEE_mLiJurhIPi13ASpoEPF2hwiV-RKNhUoXdFY')
gc = gspread.service_account(filename="mcview-starter-kit.json")
sh = gc.open_by_key('1sHhHoEE_mLiJurhIPi13ASpoEPF2hwiV-RKNhUoXdFY')

num_sensors = 0
for i in range(len(exchange_table['Name (eWon)'])) :
	if exchange_table['McView Data Historisation (YES/NO)'][i] == 'YES' :
		num_sensors = num_sensors + 1

total_column_length = int(5000000/(num_sensors * 7) + 2)
 print("total column length is :", total_column_length)

#print(respuesta.text)
pprint.pprint(respuest.json())

#j = respuest.json()


respuestjson = respuest.json()
ewonTags = respuestjson.get('ewons')[0].get('tags')
range_ewontagsinResponse = len(ewonTags)
for i in range(range_ewontagsinResponse) :
#print('ewontags', ewonTags)
	ix = respuestjson.get('ewons')[0].get('tags')[i]
	newdata = ix['history']
	dfbrut = pd.DataFrame(newdata)
	df = dfbrut.loc[:,['value', 'date']]
	df['date'] = pd.to_datetime(df.date)
	df['TagId'] = ix['ewonTagId']
	df['Hour'] = df.date.dt.hour
	df['Tagname'] = ix['name']
	df['Month'] = df.date.dt.month
	df['Extrainfo'] = exchange_table['Description (eWon)'][int(ix['ewonTagId']) - 1]
	df.rename(columns = {'value':'Value'}, inplace=True)
	df.rename(columns = {'date':'TimeStr'}, inplace=True)
	#df.rename(columns=df.iloc[0]).drop(df.index[0])
	print(df)



	#SP_PR

	#wks = spreadsheet.worksheet_by_title(ix['name'])
	wks = sh.worksheet(ix['name'])
	#worksheet = sh.worksheet("PT10")
	dataframe = pd.DataFrame(wks.get_all_records())

	# print(len(dataframe))
	# print(dataframe)

	dataframe = dataframe.append(df)
	





#read = wks.get_as_df()


#Get the data from the Sheet into python as DF

#Print the head of the datframe
#print("the length of worksheet is :", len(read))

# wks.set_dataframe(df, 'A'+str(len(read) + 2), copy_head=False)

# wks = spreadsheet.worksheet_by_title(ix['name'])
# #Get the data from the Sheet into python as DF
# read = wks.get_as_df()
# # Name (eWon)






# print("the max length of a column is :", total_column_length)

# wks.delete_rows(2, int(total_column_length/2))


# range_ewontagsinResponse = len(ewonTags)
# print(range_ewontagsinResponse)

# FLAGMoreData = respuestjson.get('moreDataAvailable')
# print(FLAGMoreData)
# for i in range(len(ewonTags)) :
# 	print('****************************************************************')
# 	ix = respuestjson.get('ewons')[0].get('tags')[i]['history'][0]['value'] # P503 ix --> store Tag X from ewon 0
# 	print(ix)
# 	print('****************************************************************')