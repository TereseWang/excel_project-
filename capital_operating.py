import xlrd
# import click
import openpyxl
# from time import strptime
# import copy
import datetime

#add and if name == main if desired

year = 18



def correct_year(year_in_question):

	'''
	This will make sure the year is correct
	Does not update it, rather, it corrects the year
	for the rest of the program.
	'''

	today = datetime.date.today()
	if(year_in_question != today.year%100):
		
		year_in_question = today.year%100
		print('updated year_in_question!')
		print(year_in_question)

		return year_in_question
	return

year = correct_year(year)

def store_values(*args):
	yr = args[0]

	#print(yr)
	'''
	This function will return a tuple of values.
	the first element of the tuple will be the dates
	and equipment costs
	the second element of the tuple will be the dates
	and software costs
	'''
	workbook = xlrd.open_workbook('Data Metrics 3 Months Trailing_07062020.xlsx')

	sheet = workbook.sheet_by_name('Capital and Operating')

	old_dates = sheet.row_values(1)	#can use search function to find location if not here
									# or for any other part where I referece a specefic row value
									#ex: store search value in cnum, pass cnum into func as a parameter and plug it into the thing
									#If you want to take this approach, I advise you read up on tuples and *args. Using *args allows the func
									#to take multiple arguments in the form of a tuple. It creates a tuple of arguments accessible in the func
	eq_money = sheet.row_values(12)

	soft_money = sheet.row_values(13)

	date_eq = {}

	date_soft = {}

	for i in range (2, len(old_dates)):

		date_eq[old_dates[i][0:3] + '-' + str(yr)] = eq_money[i]

		date_soft[old_dates[i][0:3] + '-' + str(yr)] = soft_money[i]

	return date_eq, date_soft

#print(store_values(year))

date_eq = store_values(year)[0]

date_soft = store_values(year)[1]


#Next I will use openpyxl to write the info to the outdated file