#!/usr/bin/env python
#Learn how this works here: http://youtu.be/pxofwuWTs7c

# -*- coding: utf-8 -*-
import requests
import json
import pprint
import pandas as pd
import pygsheets
import gspread
from gspread_dataframe import set_with_dataframe
import os
from time import sleep


# API Key Rungis ========> PexqYJxwfBRTG86Ae8CyxNcOeebAuQerubVuV4YFYnw9zsqjmA

exchange_table = pd.read_excel('McView Exchange Table_ HRSRungis.xlsx', sheet_name ='in', header=0)		# Exchange table

gc = gspread.service_account(filename="mcview-starter-kit.json")		# Json File for Google oauth2client Service Account Credentials
sh = gc.open_by_key('1nq_OdV7MlpdyjeAYrxx5-_3wraaD_iJLIZHH-Ra93xo')		# Open spreadsheet by key


# counting sheets number 
temp_defauts = True
num_sheets = 0
for i in range(len(exchange_table['Name GSpreadSheet'])) :
	print('index and Name GSpreadSheet :',i, exchange_table['Name GSpreadSheet'][i])
	if exchange_table['LogEnabled'][i] == 1 :
		if exchange_table['Name GSpreadSheet'][i] == 'Defauts' :
			if temp_defauts :
				num_sheets = num_sheets + 1
				temp_defauts = False
		else :
			num_sheets = num_sheets + 1

print('the number of sheets is :', num_sheets)


total_column_length = int(5000000/(num_sheets * 7) + 2)		# Calculating total coulmn length

print("total column length is :", total_column_length)
print("total column length is :", int(total_column_length/2))

# PT10 Ligne 6778 total 13479/2=6739<<<<<<<<<<<<<<<13478<<<<<<<<<<<<< 15693

while True :
	try :
		Transactionsheet = sh.worksheet('Transactions')                                  # Open tab Transactions from Speadsheet choosed
		StartingTransactionID = Transactionsheet.cell(1,2).value                # Get value from spreadSheet tab Transactions Cell (1,2)
		LastTransactionSTR = str(StartingTransactionID)

		# Sync DATA POST
		Authentication = {'t2mdevid': '87e76678-d393-4bdc-88a4-08cf164b4944','t2mtoken': 'PexqYJxwfBRTG86Ae8CyxNcOeebAuQerubVuV4YFYnw9zsqjmA', 'createTransaction':'true','lastTransactionId':LastTransactionSTR}
		respuest = requests.post('https://data.talk2m.com/syncdata', data=Authentication)
		pprint.pprint(respuest.json())
		respuestjson = respuest.json()

		TransactionID = respuestjson.get('transactionId')          # Get new tansaction ID from the ewon API request.
		LastTransactionSTR = str(TransactionID)                     # Tansaction ID to string
		print('Transaction ID: ' + str(TransactionID))
		Transactionsheet = sh.worksheet('Transactions')				# Open tab Transactions from Speadsheet choosed
		Transactionsheet.update_cell(1,2, TransactionID)			# Update transation ID 

		#Flag.. More Data Available?
		FLAGMoreData = respuestjson.get('moreDataAvailable')
		print('More Data Available: ' + str(FLAGMoreData))

		ewonTags = respuestjson.get('ewons')[0].get('tags')		# Get data from Ewon


		range_ewontagsinResponse = len(ewonTags)		# Length of the recieved package

		for i in range(range_ewontagsinResponse) :
			try :

				# Get data from Ewon and shape it into a Data frame of 7 columns
				ix = respuestjson.get('ewons')[0].get('tags')[i]		# Get tags from Ewon
				print('##############################', ix['name'], ix['ewonTagId'])
				try :
					newdata = ix['history']		# Get history of the Ewon tag 
				except :
					print("Step 1 Error : this sensor has no history")
				dfbrut = pd.DataFrame(newdata)		
				df = dfbrut.loc[:,['value', 'date']]		# Convert history into a DataFrame of two columns : value and date
				df['date'] = pd.to_datetime(df.date)
				df['TagId'] = ix['ewonTagId']
				df['Hour'] = df.date.dt.hour
				for j in range(len(exchange_table['Name GSpreadSheet'])) :
					if int(ix['ewonTagId']) == int(exchange_table['Id'][j]) :
						df['Tagname'] = exchange_table['Name'][j]

				df['Month'] = df.date.dt.month
				for j in range(len(exchange_table['Name GSpreadSheet'])) :
					if int(ix['ewonTagId']) == int(exchange_table['Id'][j]) :
						df['Extrainfo'] = exchange_table['Description'][j]
				df.rename(columns = {'value':'Value'}, inplace=True)
				df.rename(columns = {'date':'TimeStr'}, inplace=True)
				

				for j in range(len(exchange_table['Name GSpreadSheet'])) :
					if int(ix['ewonTagId']) == int(exchange_table['Id'][j]) :
						wks = sh.worksheet(str(exchange_table['Name GSpreadSheet'][j]))		# Select worsheet following the Tag Id of the recieved package
						path = 'C:/Users/admin/Desktop/McView_SKit/HRS_ROUEN_CSV_FILES/' + str(exchange_table['Name GSpreadSheet'][i]) + '.csv'		# Path to the csv file
				
				dataframe_length_measure = pd.DataFrame(wks.get_all_records())		# Get all data from the selected worksheet
				dataframe_length_measure = dataframe_length_measure.append(df)		# Length of data


				# Writing data to the selected worksheet and deleting half of it if if it exceeds the max column length
				if len(dataframe_length_measure) >= total_column_length :
					df_write_to_csv = dataframe_length_measure[:int(total_column_length/2)]		# 
					print('length dataframe to write in csv file :', len(df_write_to_csv))
					print('dataframe to write in csv file :', df_write_to_csv)
					if os.stat(path).st_size == 0:		# if csv file is empty
						df_write_to_csv.to_csv(path, sep=';' ,index=False, encoding='utf-8')
						wks.delete_rows(2, int(total_column_length/2) + 1)
						dataframe = pd.DataFrame(wks.get_all_records())
						dataframe = dataframe.append(df)
						set_with_dataframe(wks, dataframe)
						#print('dataframe is :', dataframe)
						print('df is :', df)


					else :		# if csv file isn't empty
						df_read_csv = pd.read_csv(path, sep=';', encoding='utf-8')
						dfconcat = pd.concat([df_read_csv, df_write_to_csv])
						dfconcat.to_csv(path, sep=';', index=False, encoding='utf-8')
						wks.delete_rows(2, int(total_column_length/2) + 1)
						dataframe = pd.DataFrame(wks.get_all_records())
						dataframe = dataframe.append(df)
						set_with_dataframe(wks, dataframe)
						#print('dataframe is :', dataframe)
						print('df is :', df)

				else :		# Writing data to the selected worksheet
					try :
						dataframe = pd.DataFrame(wks.get_all_records())
						dataframe = dataframe.append(df)
						set_with_dataframe(wks, dataframe)
					except :
						print("Step 2 Error : impossible to write to spread sheet")

			except :
				print("Step X Error")
			sleep(2)


		while FLAGMoreData :
			print('More data available status is :', FLAGMoreData)
			sleep(100)
			Authentication = {'t2mdevid': '87e76678-d393-4bdc-88a4-08cf164b4944','t2mtoken': 'PexqYJxwfBRTG86Ae8CyxNcOeebAuQerubVuV4YFYnw9zsqjmA', 'createTransaction':'true','lastTransactionId':LastTransactionSTR}
			respuest = requests.post('https://data.talk2m.com/syncdata', data=Authentication)
			pprint.pprint(respuest.json())
			respuestjson = respuest.json()

			TransactionID = respuestjson.get('transactionId')          # Get new Tansaction ID from the ewon API request.
			LastTransactionSTR = str(TransactionID)                     # 
			print('Transaction ID: ' + str(TransactionID))
			Transactionsheet = sh.worksheet('Transactions')
			Transactionsheet.update_cell(1,2, TransactionID)

			#Flag.. More Data Available?
			FLAGMoreData = respuestjson.get('moreDataAvailable')
			print('More Data Available: ' + str(FLAGMoreData))

			ewonTags = respuestjson.get('ewons')[0].get('tags')


			range_ewontagsinResponse = len(ewonTags)


			for i in range(range_ewontagsinResponse) :
				try :
					ix = respuestjson.get('ewons')[0].get('tags')[i]
					print('##############################', ix['name'], ix['ewonTagId'])
					try :
						newdata = ix['history']
					except :
						print("Step 1 Error : this sensor has no history")
					dfbrut = pd.DataFrame(newdata)
					df = dfbrut.loc[:,['value', 'date']]
					df['date'] = pd.to_datetime(df.date)
					df['TagId'] = ix['ewonTagId']
					df['Hour'] = df.date.dt.hour
					for j in range(len(exchange_table['Name GSpreadSheet'])) :
						if int(ix['ewonTagId']) == int(exchange_table['Id'][j]) :
							df['Tagname'] = exchange_table['Name'][j]

					df['Month'] = df.date.dt.month
					for j in range(len(exchange_table['Name GSpreadSheet'])) :
						if int(ix['ewonTagId']) == int(exchange_table['Id'][j]) :
							df['Extrainfo'] = exchange_table['Description'][j]
					df.rename(columns = {'value':'Value'}, inplace=True)
					df.rename(columns = {'date':'TimeStr'}, inplace=True)
					#print(df)
					for j in range(len(exchange_table['Name GSpreadSheet'])) :
						if int(ix['ewonTagId']) == int(exchange_table['Id'][j]) :
							wks = sh.worksheet(str(exchange_table['Name GSpreadSheet'][j]))
							path = 'C:/Users/admin/Desktop/McView_SKit/HRS_ROUEN_CSV_FILES/' + str(exchange_table['Name GSpreadSheet'][i]) + '.csv'
					
					dataframe_length_measure = pd.DataFrame(wks.get_all_records())
					dataframe_length_measure = dataframe_length_measure.append(df)

					if len(dataframe_length_measure) >= total_column_length :
						df_write_to_csv = dataframe_length_measure[:int(total_column_length/2)]
						print('length dataframe to write in csv file :', len(df_write_to_csv))
						print('dataframe to write in csv file :', df_write_to_csv)
						if os.stat(path).st_size == 0:
							df_write_to_csv.to_csv(path, sep=';' ,index=False, encoding='utf-8')
							wks.delete_rows(2, int(total_column_length/2) + 1)
							dataframe = pd.DataFrame(wks.get_all_records())
							dataframe = dataframe.append(df)
							set_with_dataframe(wks, dataframe)
							#print('dataframe is :', dataframe)
							print('df is :', df)


						else :
							df_read_csv = pd.read_csv(path, sep=';', encoding='utf-8')
							dfconcat = pd.concat([df_read_csv, df_write_to_csv])
							dfconcat.to_csv(path, sep=';', index=False, encoding='utf-8')
							wks.delete_rows(2, int(total_column_length/2) + 1)
							dataframe = pd.DataFrame(wks.get_all_records())
							dataframe = dataframe.append(df)
							set_with_dataframe(wks, dataframe)
							#print('dataframe is :', dataframe)
							print('df is :', df)

					else :
						try :
							dataframe = pd.DataFrame(wks.get_all_records())
							dataframe = dataframe.append(df)
							set_with_dataframe(wks, dataframe)
						except :
							print("Step 2 Error : impossible to write to spread sheet")

				except :
					print("Step X Error")
				sleep(2)


	except :
		print("Step XX Error")