import pygsheets
import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile
import numpy as np
import csv


def creat_sheets(name) :
	worksheet = spreadsheet.add_worksheet(title=name, cols="7")

def manipulate_spsh(name) :
	list = [['Value','TimeStr','TagId','Hour','Tagname','Month','Extrainfo'],['0','2021-05-19 14:09:00+00:00','194','14','5','19','PT x']]
	wks = spreadsheet.worksheet_by_title(name)
	wks.update_values('A1:G2', values = list)


gc = pygsheets.authorize(service_file='mcview-starter-kit.json')
spreadsheet = gc.open_by_key('1nq_OdV7MlpdyjeAYrxx5-_3wraaD_iJLIZHH-Ra93xo')
df = pd.read_excel('McView Exchange Table_ HRSRungis.xlsx', sheet_name ='in', header=0)

worksheet = spreadsheet.add_worksheet('Transactions',rows="9", cols="3")
wks = spreadsheet.worksheet_by_title('Transactions')
wks.update_value('A1', 'Transaction ID')

temp_defauts = True

for i in range(len(df['Name GSpreadSheet'])) :
	if df['LogEnabled'][i] == 1 :
		if df['Name GSpreadSheet'][i] == 'Defauts' :
			if temp_defauts :
				creat_sheets(df['Name GSpreadSheet'][i])
				temp_defauts = False
		else :
			creat_sheets(df['Name GSpreadSheet'][i])


for i in range(len(df['Name GSpreadSheet'])) :
	if df['LogEnabled'][i] == 1 :
		if df['Name GSpreadSheet'][i] == 'Defauts' :
			if not temp_defauts :
				manipulate_spsh(df['Name GSpreadSheet'][i])
				temp_defauts = True
		else :
			manipulate_spsh(df['Name GSpreadSheet'][i])

# for i in range(len(df['Name (eWon)'])) :
# 	if df['McView Data Historisation (YES/NO)'][i] == 'YES' :
# 		manipulate_spsh(df['Name (eWon)'][i])



# df_drop = df.drop_duplicates('Name GSpreadSheet')
# indexNames = df_drop[ df_drop['LogEnabled'] == 0 ].index
# df_drop.drop(indexNames , inplace=True)
# print(df_drop)
# list_sheets = df_drop['Name GSpreadSheet']



# for i in range(len(list_sheets)) :
# 	creat_sheets(str(list_sheets[i]))


# for i in range(len(list_sheets)) :
# 	manipulate_spsh(str(list_sheets[i]))

# for i in range(len(df['Name (eWon)'])) :
# 	if df['McView Data Historisation (YES/NO)'][i] == 'YES' :
# 		creat_sheets(df['Name (eWon)'][i])


# for i in range(len(df['Name (eWon)'])) :
# 	if df['McView Data Historisation (YES/NO)'][i] == 'YES' :
# 		manipulate_spsh(df['Name (eWon)'][i])


#df = pd.read_excel('McView Exchange Table_ HRSRouen.xlsx', sheet_name ='in', header=0)
# for i in range(len(df['Name (eWon)'])) :
# 	if df['McView Data Historisation (YES/NO)'][i] == 'YES' :
# 		path = 'C:/Users/admin/Desktop/McView_SKit/HRS_ROUEN_CSV_FILES/' + str(df['Name (eWon)'][i]) + '.csv'
# 		with open(path, 'w', encoding='UTF8', newline='') as f:
# 			writer = csv.writer(f)