import pygsheets
import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile
import numpy as np
from threading import Thread
import threading
from time import sleep


def creat_sheets(name) :
	worksheet = spreadsheet.add_worksheet(title=name, cols="7")

def manipulate_spsh(name) :
	list = [['Value','TimeStr','TagId','Hour','Tagname','Month','Extrainfo'],['0','2021-05-19 14:09:00+00:00','194','14','5','19','PT x']]
	wks = spreadsheet.worksheet_by_title(name)
	wks.update_values('A1:G2', values = list)


gc = pygsheets.authorize(service_file='mcview-starter-kit.json')
spreadsheet = gc.open_by_key('1gECWnECgIoTQ3YGT5QZNwTsvQMKz5AymReyxX7P-XMI')
df = pd.read_excel('GUS_SMTAG.xlsx', sheet_name ='var_lst', header=0)

worksheet = spreadsheet.add_worksheet('Transactions',rows="9", cols="3")
wks = spreadsheet.worksheet_by_title('Transactions')
wks.update_value('A1', 'Transaction ID')

for i in range(len(df['Name'])) :
	if df['LogEnabled'][i] == 1 :
		creat_sheets(df['Name'][i])


for i in range(len(df['Name'])) :
	if df['LogEnabled'][i] == 1 :
		manipulate_spsh(df['Name'][i])