import pandas as pd
import datetime
import warnings
from googletrans import Translator
from datetime import datetime

# This removes the default header style so we can override it later
import pandas.io.formats.excel
pandas.io.formats.excel.header_style = None

#Parameter:
START_DATE = '2019-05-01'
dt = datetime.strptime(START_DATE, '%Y-%m-%d')
M = dt.strftime('%B')
END_DATE = '2019-05-10'
PATH = './May_01_May_10.csv'
#Output the Top N expense
TOP_N = 10
#For graph purpose
SHOWN = 15
SAVE_PATH = '/Users/WhatEverPathUWant/'

#Real Function starts from here
def expense():
	#FORMATTING:
	warnings.filterwarnings("ignore")
	pd.set_option("display.max_rows", 120)
	pd.set_option("display.max_columns",10)
	month = pd.read_csv(PATH)
	month['Date'] = pd.to_datetime(month.Date)

	#Mask the timeframe between the start and end date you choose
	mask = (month['Date'] > START_DATE) & (month['Date'] <= END_DATE)
	df = month.loc[mask]
	k = 0 
	for i in df['Amount'].keys():
		for j in "($,)":
			df['Amount'][i] = df['Amount'][i].replace(j,"")
		df['Amount'][i] = float(df['Amount'][i])
	df = df.sort_values(by = ['Amount'], ascending = False)
	blankIndex = [''] * len(df)
	df.index = blankIndex
	exp = df[['Description','Category','Amount']]

	# Translate the description row and category row by importing Google API
	exp["花费描述谷歌翻译"] = ""
	exp["类别"] = ""
	for index, row in exp.iterrows():
			translator = Translator()
			chi_text = translator.translate(row['Description'], src ="en", dest = "zh-cn")
			chi_cat = translator.translate(row['Category'], src ="en", dest = "zh-cn")
			row["花费描述谷歌翻译"] = chi_text.text
			row["类别"] = chi_cat.text

	exp = exp[['Description','花费描述谷歌翻译','Category','类别','Amount']]

	acct_cat = df[['Account Name','Amount']]
	acct_cat = acct_cat.groupby(['Account Name']).sum() .reset_index()
	acct_cat = acct_cat.sort_values(by = ['Amount'], ascending = False)
	print('\n')
	print("Top spent by Bank Account:")
	print(acct_cat['Account Name'])
	print('\n')
	translator = Translator()
	acct_cat['银行账户名称'] = acct_cat['Account Name'].map(lambda x: translator.translate(x,src = "en", dest = "zh-cn").text)
	print(acct_cat)
	acct_cat = acct_cat[['Account Name','银行账户名称','Amount']]

	blankIndex = [''] * len(acct_cat)
	acct_cat.index = blankIndex	
	print(acct_cat)
	desc_cat = df[['Description','Amount']]
	desc_cat = desc_cat.groupby(['Description']).sum().reset_index()
	desc_cat = desc_cat.sort_values(by = ['Amount'], ascending = False)
	blankIndex = [''] * len(desc_cat)
	desc_cat.index = blankIndex
	print('\n')
	print("Top spent by Description:")
	print(desc_cat)
	print('\n')

	cat_cat = df[['Category','Amount']]
	cat_cat = cat_cat.groupby(['Category']).sum().reset_index()
	cat_cat = cat_cat.sort_values(by = ['Amount'], ascending = False)
	blankIndex = [''] * len(cat_cat)
	cat_cat.index = blankIndex
	cat_cat['支出分类'] = cat_cat['Category'].map(lambda x: translator.translate(x,src = "en", dest = "zh-cn").text)
	cat_cat = cat_cat[['Category','支出分类','Amount']]
	print("Top spent by Category:")
	print(cat_cat)
	print('\n')

	#Print out the spending summary
	print('你这段时间最大的 %d 项花费一共多少钱$ :' %(TOP_N), exp['Amount'][:TOP_N].sum())	
	print("从 %s 到 %s 你总共花销是$:" %(START_DATE, END_DATE), df['Amount'].sum())

	#Return the dataframes
	return df, exp, acct_cat, cat_cat

def output(df, exp, acct_cat, cat_cat):

	# Create Format objects to apply to sheet
	# https://xlsxwriter.readthedocs.io/format.html#format-methods-and-format-properties
	START_STRIP = START_DATE.replace('-','')
	END_STRIP = END_DATE.replace('-','')
	sheet = '%s到%s花销'%(START_STRIP, END_STRIP)
	writer = pd.ExcelWriter('%s%s.xlsx' %(SAVE_PATH,M), engine = 'xlsxwriter')

	#Appending all the dataframes to one:
	df_list = []
	df_list.append(exp)
	df_list.append(acct_cat)
	df_list.append(cat_cat)

	row = 3
	space = 2 

	#Print all dataframe to same excel
	for dataframe in df_list:
		dataframe.to_excel(writer, sheet, startrow = row, startcol = 0, index = False)
		row = row + len(dataframe.index) + space

	#Creating workbook and worksheet using xlsxwriter
	workbook = writer.book
	worksheet = writer.sheets[sheet]

	#Insert Charts
	#Reset row to 3, space =2, for the chart variables
	row = 3 
	space = 2 
	chart = workbook.add_chart({'type': 'column'})
	acct_row = len(acct_cat.index)
	top_exp = exp[:TOP_N]
	top_expr = len(top_exp.index)
	exp_row = len(exp.index)
	cat_row = len(cat_cat.index)

	#Add charts
	chart.add_series({
		'name':	 '账户使用数据',
		'categories': '=%s!$B$%d:$B$%d' %(sheet, (row+2), (row + 2 + SHOWN)),
		'data_labels': {'value': True},
		'values': '=%s!$E$%d:$E$%d' % (sheet, (row + 2), (row + 2 + SHOWN))})
	chart.set_title({'name': '账户使用数据'})
	chart.set_x_axis({'name': '消费描述', 'text_axis': True})
	worksheet.insert_chart('G1', chart,{'x_scale': 2, 'y_scale': 2})

	chart1 = workbook.add_chart({'type': 'pie'})
	chart1.add_series({
		'name':	 '账户使用数据',
		'categories': '=%s!$B$%d:$B$%d' % (sheet, (row+2 ), (row + 2 + SHOWN)),
		'data_labels': {'value': True},
		'values': '=%s!$E$%d:$E$%d' % (sheet, (row+2), (row + 2 + SHOWN))})
	chart1.set_title({'name': '账户使用数据'})
	worksheet.insert_chart('G33', chart1, {'x_scale': 1.7, 'y_scale': 1.7})

	chart3 = workbook.add_chart({'type': 'column'})
	chart3.add_series({
		'name':	 '银行账户数据',
		'categories': '=%s!$B$%d:$B$%d' %(sheet, (exp_row +row+2 + space), (exp_row + acct_row + 3 + row)),
		'data_labels': {'value': True},
		'values': '=%s!$C$%d:$C$%d' % (sheet, (exp_row +row+2 + space), (exp_row + acct_row + 3 + row))})
	chart3.set_title({'name': '银行账户数据'})
	chart3.set_x_axis({'name': '账户名称', 'text_axis': True})
	worksheet.insert_chart('E%d' % (exp_row + row + space + 1), chart3,{'x_scale': 1.5, 'y_scale': 1.5})

	chart4 = workbook.add_chart({'type': 'pie'})
	chart4.add_series({
		'name':	 '银行账户数据',
		'categories': '=%s!$B$%d:$B$%d' % (sheet, (exp_row +row+2 + space), (exp_row + acct_row + 3 + row)),
		'data_labels': {'value': True},
		'values': '=%s!$C$%d:$C$%d' % (sheet, (exp_row +row+2 + space), (exp_row + acct_row + 3 + row))})
	chart4.set_title({'name': '银行账户数据'})
	worksheet.insert_chart('Q%d' % (exp_row + row + space + 1), chart4, {'x_scale': 1.7, 'y_scale': 1.5})

	chart5 = workbook.add_chart({'type': 'column'})
	chart5.add_series({
		'name':	 '类别使用数据',
		'categories': '=%s!$B$%d:$B$%d' %(sheet, (exp_row + acct_row + 4 + row + space), (exp_row + acct_row + cat_row + 5 + row)),
		'data_labels': {'value': True},
		'values': '=%s!$C$%d:$C$%d' % (sheet, (exp_row + acct_row + 4 + row + space), (exp_row + acct_row + cat_row + 5 + row))})
	chart5.set_title({'name': '类别使用数据'})
	chart5.set_x_axis({'name': '类别名', 'text_axis': True})
	worksheet.insert_chart('A%d' % (exp_row + acct_row + cat_row + space*2 + row + 5), chart5,{'x_scale': 1.8, 'y_scale': 1})

	chart6 = workbook.add_chart({'type': 'pie'})
	chart6.add_series({
		'name':	 '类别使用数据',
		'categories': '=%s!$B$%d:$B$%d' % (sheet, (exp_row + acct_row + 4 + row + space), (exp_row + acct_row + cat_row + 5 + row)),
		'data_labels': {'value': True},
		'values': '=%s!$C$%d:$C$%d' % (sheet, (exp_row + acct_row + 4 + row + space), (exp_row + acct_row + cat_row + 5 + row))})
	chart6.set_title({'name': '类别使用数据'})
	worksheet.insert_chart('E%d' % (exp_row + acct_row + cat_row + space*2 + row + 5), chart6,{'x_scale': 1.7, 'y_scale': 2})

	#Apply formatting to sheet
	red_bold = workbook.add_format({'bold': True, 'font_color': 'red'})
	digit = workbook.add_format({'num_format': '#,##0.00'})
	text_wrap = workbook.add_format({
	    'bold': True,
	    'text_wrap': True
	    })
	worksheet.set_column('A:A', None, text_wrap)
	worksheet.set_column('B:B', None, text_wrap)
	worksheet.set_column('C:C', None, text_wrap)
	worksheet.set_column('E:E', None, digit)
	worksheet.set_column('A:A',50)
	worksheet.set_column('B:B',50)
	worksheet.set_column('C:C',20)
	worksheet.set_column('D:D',20)

	# Apply a conditional format to a cell range.
	worksheet.conditional_format('E2:E%d'%(exp_row + 1 + row), {'type': '3_color_scale'})
	worksheet.conditional_format('C%d:C%d' %((exp_row + 2 + row),(exp_row + acct_row + 3 + row)), {'type': '3_color_scale'})
	worksheet.conditional_format('C%d:C%d'%((exp_row + acct_row + 4 + row), (exp_row + acct_row + cat_row + 5 + row)), {'type': '3_color_scale'})

	#Save File
	worksheet.write(0, 0 , '你这段时间最大的 %d 项花费一共多少钱$ %d' %(TOP_N, exp['Amount'][:TOP_N].sum()))
	worksheet.write(1, 0 , "从 %s 到 %s 你总共花销是$ %d" %(START_DATE, END_DATE, df['Amount'].sum()))
	writer.save()

#Call Function and Print out
df, exp, acct_cat, cat_cat = expense()
output(df, exp, acct_cat, cat_cat)



