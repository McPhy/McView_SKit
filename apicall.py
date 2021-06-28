# #!/usr/bin/env python
# #Learn how this works here: http://youtu.be/pxofwuWTs7c
# # import requests
# # import json
# # import pprint
import pygsheets
import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile
import numpy as np

gc = pygsheets.authorize(service_file='mcview-starter-kit.json')
# res = gc.sheet.create("HRS-ROUEN 2.0")  # Please set the new Spreadsheet name.
# print(res)
spreadsheet = gc.open_by_key('1M80r-50kEP5vWWhKSJlog7BrHl7T24glcPR19DwOJ-I')
df = pd.read_excel('McView Exchange Table_ HRSRungis.xlsx', sheet_name ='McView Exchange Table_ HRSRungi', header=0)

worksheet = spreadsheet.add_worksheet('Transactions',rows="9", cols="3")
# worksheet = spreadsheet.open('HRS-RUNGIS')
wks = spreadsheet.worksheet_by_title('Transactions')
wks.update_value('A1', 'Transaction ID')

for i in range(len(df['Name'])) :
	if df['McView Data Historisation (YES/NO)'][i] == 'YES' :
		worksheet = spreadsheet.add_worksheet(title=df['Name'][i], cols="7")
		# worksheet = spreadsheet.open('HRS-RUNGIS')
		wks = spreadsheet.worksheet_by_title(df['Name'][i])
		# value = worksheet.cell('A1')
		# timestr = worksheet.cell('B1')
		# tagid = worksheet.cell('C1')
		# hour = worksheet.cell('D1')
		# tagname = worksheet.cell('E1')
		# month = worksheet.cell('F1')
		# extrainfo = worksheet.cell('G1')
		wks.update_value('A1', 'Value')
		wks.update_value('B1', 'TimeStr')
		wks.update_value('C1', 'TagId')
		wks.update_value('D1', 'Hour')
		wks.update_value('E1', 'Tagname')
		wks.update_value('F1', 'Month')
		wks.update_value('G1', 'Extrainfo')

		wks.update_value('A2', '0')
		wks.update_value('B2', '2021-05-19 14:09:00+00:00')
		wks.update_value('C2', '194')
		wks.update_value('D2', '14')
		wks.update_value('E2', '5')
		wks.update_value('F2', '19')
		wks.update_value('G2', 'PT x')


#0	2021-05-19 14:09:00+00:00	194	14	5	19	PT x
# createdSpreadsheet.share('mcphy.skit@gmail.com', role='writer', type='user')
#worksheet = createdSpreadsheet.add_worksheet(title="A worksheet", cols="10")