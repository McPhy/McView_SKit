#This code is used to create a CSV files for for every stored variable

import pygsheets
import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile
import numpy as np
import csv

df = pd.read_excel('McView Exchange Table_ HRSRungis.xlsx', sheet_name ='in', header=0)

temp_defauts = True

for i in range(len(df['Name'])) :
	if df['LogEnabled'][i] == 1 :
		if df['Name GSpreadSheet'][i] == 'Defauts' :
			if temp_defauts :
				path = 'C:/Users/admin/Desktop/McView_SKit/HRS_ROUEN_CSV_FILES/' + str(df['Name GSpreadSheet'][i]) + '.csv'
				with open(path, 'w', encoding='UTF8', newline='') as f:
					writer = csv.writer(f)
				temp_defauts = False
		else :
			path = 'C:/Users/admin/Desktop/McView_SKit/HRS_ROUEN_CSV_FILES/' + str(df['Name GSpreadSheet'][i]) + '.csv'
			with open(path, 'w', encoding='UTF8', newline='') as f:
				writer = csv.writer(f)