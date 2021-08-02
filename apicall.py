"""
This code is used to create worksheets, fill the headers and the first line with the following arrays :
['Value','TimeStr','TagId','Hour','Tagname','Month','Extrainfo']
['0','2021-05-19 14:09:00+00:00','194','14','5','19','PT x']
"""

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