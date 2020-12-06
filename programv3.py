#!/usr/bin/python3
import matplotlib.pyplot as plt 
import sqlite3 as sql 
import re 
import xlrd as xl
import os
import csv


if os.path.exists("database.db"):
	os.remove("database.db")
	pass

con = sql.connect("database.db") 															#connecting to the DB
cur = con.cursor()																			#table for the data inside the DB
cur.execute("""CREATE TABLE IF NOT EXISTS Arrivals(Year int NOT NULL, 						
												   Quarter int NOT NULL,
												   Country VARCHAR(255) NOT NULL,
												   Plane int,
												   Train int,
												   Ship int,
												   Car int,
												   Num int,
												   PRIMARY KEY (Year,Quarter,Country));""")

if not os.path.exists("excels"):															#change directory to find excel files
	print("Something went wrong with your download_script.py")
	exit(0)
os.chdir("excels")	
if not os.path.exists("../csv_files"):														#create the csv file if it does not exist
	os.mkdir("../csv_files")																
for x in range(2011,2016):																	#for each year we're interested in
	
	workbook_name = "A2001_STO04_TB_QQ_04_"+str(x)+"_02_F_GR.xls"							#assigning new workbook
	workbook = xl.open_workbook(workbook_name)												#opening it
	index = 2																				#var to keep track of quarters
	prev = 0																				#var to keep track of start of quarter
	while(index <= 11):																		#11=months
		sheet = workbook.sheet_by_index(index)
		
		prevsheet = workbook.sheet_by_index(prev)
		
		for z in range(74,sheet.nrows):														#looking for start of second table data
			if sheet.cell(z,1).value == u'Αυστρία':
				start = z
				break
		
		for z in range(74,prevsheet.nrows):													#looking for start from previous sheet
			if prevsheet.cell(z,1).value == u'Αυστρία':
				start2 = z
		
		gap = start2-start																	#accounting for human error
		
		for row in range(start,sheet.nrows - 1):
			
			if sheet.cell(row,1).value == u'ΓΕΝΙΚΟ ΣΥΝΟΛΟ': 								#reached end of table
				break
			if sheet.cell(row,1).value == '' or sheet.cell(row,1).value == u'από τίς οποίες:': #non-important data
				continue
			if index == 2:																	#if its march just enter it
				Country = str(sheet.cell(row,1).value)
				Plane = sheet.cell(row,2).value
				Train = sheet.cell(row,3).value
				Ship = sheet.cell(row,4).value
				Car = sheet.cell(row,5).value
				Arrivals = sheet.cell(row,6).value
			if x==2013 and index == 8 and row == 89:
					statement = "INSERT INTO Arrivals VALUES(%d,%d,'%s',%d,%d,%d,%d,%d)"%(x,index+1,u'Κροατία',1152,0,2534,0,3686) #formatting query
					cur.execute(statement)
					gap-=1
					
					continue
			elif gap==0: 																	#if it's not march and there isn't a gap 
																							#sub from same cell in previous sheet
																							#(rest is same but accounting for differences in height of corresponding cells)
				if x == 2013 and index==11 and row == 89:
					statement = "INSERT INTO Arrivals VALUES(%d,%d,'%s',%d,%d,%d,%d,%d)"%(x,index+1,u'Κροατία',1360,0,3021,0,4381) #formatting query
					cur.execute(statement)
					
					
					continue
				Country = str(sheet.cell(row,1).value)
				Plane = round(sheet.cell(row,2).value) - round(prevsheet.cell(row+gap,2).value)
				Train = round(sheet.cell(row,3).value) - round(prevsheet.cell(row+gap,3).value)
				Ship = round(sheet.cell(row,4).value) - round(prevsheet.cell(row+gap,4).value)
				Car = round(sheet.cell(row,5).value) - round(prevsheet.cell(row+gap,5).value)
				Arrivals = round(sheet.cell(row,6).value) - round(prevsheet.cell(row+gap,6).value)
				
			elif gap > 0:
				
				Country = str(sheet.cell(row,1).value)
				Plane = round(sheet.cell(row,2).value) - round(prevsheet.cell(row+gap,2).value)
				Train = round(sheet.cell(row,3).value) - round(prevsheet.cell(row+gap,3).value)
				Ship = round(sheet.cell(row,4).value) - round(prevsheet.cell(row+gap,4).value)
				Car = round(sheet.cell(row,5).value) - round(prevsheet.cell(row+gap,5).value)
				Arrivals = round(sheet.cell(row,6).value) - round(prevsheet.cell(row+gap,6).value)
			elif gap < 0:

				Country = str(sheet.cell(row,1).value)
				Plane = round(sheet.cell(row,2).value) - round(prevsheet.cell(row-abs(gap),2).value)
				Train = round(sheet.cell(row,3).value) - round(prevsheet.cell(row-abs(gap),3).value)
				Ship = round(sheet.cell(row,4).value) - round(prevsheet.cell(row-abs(gap),4).value)
				Car = round(sheet.cell(row,5).value) - round(prevsheet.cell(row-abs(gap),5).value)
				Arrivals = round(sheet.cell(row,6).value) - round(prevsheet.cell(row-abs(gap),6).value)
			
			statement = "INSERT INTO Arrivals VALUES(%d,%d,'%s',%d,%d,%d,%d,%d)"%(x,index+1,Country,Plane,Train,Ship,Car,Arrivals) #formatting query
			cur.execute(statement)
		if index==2: #updating sheet "pointers"
			prev+=2
		else:
			prev +=3
		index +=3
		
		
		
con.commit() #commit as to not lose data

for row in cur.execute('SELECT * FROM Arrivals'): #printing all db contents, can be ommited
        print(row)


finlist = []
for x in range(2011,2016):																	#for each year we're interested in
	workbook_name = "A2001_STO04_TB_QQ_04_"+str(x)+"_02_F_GR.xls"							#assigning new workbook
	workbook = xl.open_workbook(workbook_name)
	sheet = workbook.sheet_by_index(11)
	for pos_row in range(134,sheet.nrows):
		if sheet.cell(pos_row,1).value == u'ΓΕΝΙΚΟ ΣΥΝΟΛΟ':								#checking for ΓΕΝΙΚΟ ΣΥΝΟΛΟ and getting the value
			finlist.append(round(sheet.cell(pos_row,6).value))
years = [i for i in range(2011,2016)]
plt.bar(years,finlist)																	#plotting it with a bar graph
plt.ylim(10000000,24000000)
plt.ylabel("Visitors")
plt.xlabel("Years")			#Making it pretty
plt.title("Visitors Each Year")
plt.show()



fig,ax = plt.subplots(nrows=3,ncols=2) #Second Plot,Getting an "array" of subplots
rows = 0
cols = 0								#Vars to help with subplot structure
fincountrylist = [] 					#Vars for csv's
finvisitslist = []								
for x in range(2011,2016):
	workbook_name = "A2001_STO04_TB_QQ_04_"+str(x)+"_02_F_GR.xls"							#assigning new workbook
	workbook = xl.open_workbook(workbook_name)
	sheet = workbook.sheet_by_index(11)
	countries_and_visits = []
	countries = []
	visits = []
	for z in range(74,sheet.nrows):														#looking for start of second table data
			if sheet.cell(z,1).value == u'Αυστρία':
				start = z
				
				break
	for row in range(start,sheet.nrows-1):

		if sheet.cell(row,1).value==u'από τίς οποίες:' or sheet.cell(row,1).value =='':
			continue
		elif sheet.cell(row,1).value == u'ΓΕΝΙΚΟ ΣΥΝΟΛΟ':
			break
		countries_and_visits.append([str(sheet.cell(row,1).value),round(sheet.cell(row,6).value)]) #append all countries with the people who 
																								   #came from there

	sorted_final = sorted(countries_and_visits,key = lambda x: x[1], reverse = True)
	for top in range(10):																		#sorting to keep top 10 and plotting them
		countries.append(sorted_final[top][0])
		visits.append(sorted_final[top][1])

	ax[rows][cols].bar(countries,visits)
	ax[rows][cols].set_title("Year = %d"%(x))
	ax[rows][cols].tick_params(axis='x',which='major',labelsize=7,labelrotation=45) #Fixing name overlap
	#ax[rows][cols].set_xlabel("countries")
	ax[rows][cols].set_ylabel("Visitors")
	fincountrylist.append(countries)
	finvisitslist.append(visits)
	if cols==1:
		rows+=1 #Custom array iterator,Rest of plots are the same thing
		cols=0
		continue
	cols+=1
	
plt.tight_layout(pad=2)
plt.show()


transport = ['Plane','Train','Ship','Car']
rows = 0
cols = 0
fig,ax = plt.subplots(nrows=3,ncols=2)
finvisitorslist = []
for x in range(2011,2016):																	#for each year we're interested in
	workbook_name = "A2001_STO04_TB_QQ_04_"+str(x)+"_02_F_GR.xls"							#assigning new workbook
	workbook = xl.open_workbook(workbook_name)
	sheet = workbook.sheet_by_index(11)
	for pos_row in range(134,sheet.nrows):
		if sheet.cell(pos_row,1).value == u'ΓΕΝΙΚΟ ΣΥΝΟΛΟ':								#checking for ΓΕΝΙΚΟ ΣΥΝΟΛΟ and getting the values
			
			visitors = [round(sheet.cell(pos_row,2).value),
						round(sheet.cell(pos_row,3).value),
						round(sheet.cell(pos_row,4).value),
						round(sheet.cell(pos_row,5).value)]



	ax[rows][cols].bar(transport,visitors)
	ax[rows][cols].set_title("Year = %d"%(x))
	ax[rows][cols].tick_params(axis='x',which='major',labelsize=7)
	ax[rows][cols].set_xlabel("Means of Transport")
	ax[rows][cols].set_ylabel("Visitors")
	finvisitorslist.append(visitors)
	if cols==1:
		rows+=1
		cols=0
		continue
	cols+=1
plt.tight_layout(pad=2)
plt.ylabel("visitors")
plt.xlabel("Means of Transport")
plt.show()


Quarters = ['Q1','Q2','Q3','Q4']
rows = 0
cols = 0
fig,ax = plt.subplots(nrows=3,ncols=2)
finYQlist = []

for x in range(2011,2016):																	#for each year we're interested in
	Visits = []
	
	workbook_name = "A2001_STO04_TB_QQ_04_"+str(x)+"_02_F_GR.xls"							#assigning new workbook
	workbook = xl.open_workbook(workbook_name)												#opening it
	index = 2																				#var to keep track of quarters
	prev = 0																				#var to keep track of start of quarter
	while(index <= 11):																		#11=months
		sheet = workbook.sheet_by_index(index)												#current sheet
		
		prevsheet = workbook.sheet_by_index(prev)											# previous sheet
		
		for row in range(134,sheet.nrows-1):												#try to find the row with ΓΕΝΙΚΟ ΣΥΝΟΛΟ in 
																							#current sheet and keep value
			
			if sheet.cell(row,1).value==u'ΓΕΝΙΚΟ ΣΥΝΟΛΟ':
				end = round(sheet.cell(row,6).value)
				break
		for row in range(134,prevsheet.nrows-1):											#same but for previous sheet
			
			if prevsheet.cell(row,1).value==u'ΓΕΝΙΚΟ ΣΥΝΟΛΟ':
				start = round(prevsheet.cell(row,6).value)
				break
		Visits.append(end-start)															#appending diff to get the value for a quarter
		if index==2: #updating sheet "pointers"
			prev+=2
		else:
			prev +=3
		index +=3
	
	ax[rows][cols].bar(Quarters,Visits)														#plotting
	ax[rows][cols].set_title("Year = %d"%(x))
	ax[rows][cols].tick_params(axis='x',which='major',labelsize=7)
	ax[rows][cols].set_xlabel("Quarters")
	ax[rows][cols].set_ylabel("Visitors")
	finYQlist.append(Visits)
	if cols==1:
		rows+=1
		cols=0
		continue
	cols+=1
plt.tight_layout(pad=2)

plt.show()


os.chdir("../csv_files")										#changing directories to put the csv's
with open("Sum_by_year.csv",'w',newline='') as f:				#open the file I will use
	writer = csv.writer(f)										#use a writer from the csv module
	writer.writerow(['2011','2012','2013','2014','2015'])		#use the writerow() func to automatically configure spacing
	writer.writerow(finlist)									#write the rest of the list of sum values
f.close()

with open("Top_Countries_by_Year.csv",'w',newline='') as f:    #same as above
	writer = csv.writer(f)
	writer.writerow(['Year','Country','Number'])
	
	for x in range(2011,2016):
		for y in range(10):
			writer.writerow([str(x),fincountrylist[x%2010-1][y],finvisitslist[x%10-1][y]])  #kept a list of lists of each year containing the data 
f.close()																					#x%2010-1 is used to access the list based on the year first 0 second 1 and so on

with open("Visits_per_Transport_Per_Year.csv",'w',newline='') as f:
	writer = csv.writer(f)
	writer.writerow(['Year','Plane','Train','Ship','Car'])
	for x in range(2011,2016):
		writer.writerow([str(x),finvisitorslist[x%10-1][0],finvisitorslist[x%10-1][1],finvisitorslist[x%10-1][2],finvisitorslist[x%10-1][3]])
f.close()

with open("Visits_by_Year_and_Quarter.csv",'w',newline='') as f:
	writer = csv.writer(f)
	writer.writerow(['Year','Quarter','Number'])
	for x in range(2011,2016):
		for y in range(4):
			writer.writerow([str(x),str(y+1),finYQlist[x%10-1][y]])
f.close()
