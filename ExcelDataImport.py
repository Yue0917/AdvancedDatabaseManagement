import xlrd
import MySQLdb

# Open the workbook and define the worksheet
book = xlrd.open_workbook("globalterrorismdb_0617dist 2.xlsx")
sheet = book.sheet_by_name("Data")

# Establish a MySQL connection
database = MySQLdb.connect ("127.0.0.1","root","Wanqing0930","Terrorism", use_unicode=True, charset="utf8")

# Get the cursor, which is used to traverse the database, line by line
cursor = database.cursor()

# Create the INSERT INTO sql query
query = """INSERT INTO globalterrorism (eventid,iyear,imonth,iday,extended,country,country_txt,region,region_txt,provstate,city,vicinity,location,attacktype1,attacktype1_txt,targtype1,targtype1_txt,targsubtype1,targsubtype1_txt,corp1,target1,natlty1,natlty1_txt,gname,guncertain1,individual,weaptype1,weaptype1_txt,weapsubtype1,weapsubtype1_txt,propextent,propextent_txt,propvalue,dbsource,INT_LOG,INT_IDEO,INT_MISC,INT_ANY) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)"""
print('Start')
# Create a For loop to iterate through each row in the XLS file, starting at row 2 to skip the headers
for r in range(1, sheet.nrows):
		eventid		= sheet.cell(r,0).value
		iyear	        = sheet.cell(r,1).value
		imonth		= sheet.cell(r,2).value
		iday		= sheet.cell(r,3).value
		extended    	= sheet.cell(r,4).value
		country	        = sheet.cell(r,5).value
		country_txt 	= sheet.cell(r,6).value
		region		= sheet.cell(r,7).value
		region_txt  	= sheet.cell(r,8).value
		provstate   	= sheet.cell(r,9).value
		city		= sheet.cell(r,10).value
		vicinity	= sheet.cell(r,11).value
		location	= sheet.cell(r,12).value
		attacktype1     = sheet.cell(r,13).value
		attacktype1_txt = sheet.cell(r,14).value
		targtype1       = sheet.cell(r,15).value
		targtype1_txt   = sheet.cell(r,16).value
		targsubtype1    = sheet.cell(r,17).value
		targsubtype1_txt= sheet.cell(r,18).value
		corp1           = sheet.cell(r,19).value
		target1         = sheet.cell(r,20).value
		natlty1         = sheet.cell(r,21).value
		natlty1_txt     = sheet.cell(r,22).value
		gname           = sheet.cell(r,23).value
		guncertain1     = sheet.cell(r,24).value
		individual      = sheet.cell(r,25).value
		weaptype1       = sheet.cell(r,26).value
		weaptype1_txt   = sheet.cell(r,27).value
		weapsubtype1    = sheet.cell(r,28).value
		weapsubtype1_txt= sheet.cell(r,29).value
		propextent      = sheet.cell(r,30).value
		propextent_txt  = sheet.cell(r,31).value
		propvalue       = sheet.cell(r,32).value
		dbsource        = sheet.cell(r,33).value
		INT_LOG         = sheet.cell(r,34).value
		INT_IDEO        = sheet.cell(r,35).value
		INT_MISC        = sheet.cell(r,36).value
		INT_ANY         = sheet.cell(r,37).value
		

		# Assign values from each row
		values = (eventid,iyear,imonth,iday,extended,country,country_txt,region,region_txt,provstate,city,vicinity,location,attacktype1,attacktype1_txt,targtype1,targtype1_txt,targsubtype1,targsubtype1_txt,corp1,target1,natlty1,natlty1_txt,gname,guncertain1,individual,weaptype1,weaptype1_txt,weapsubtype1,weapsubtype1_txt,propextent,propextent_txt,propvalue,dbsource,INT_LOG,INT_IDEO,INT_MISC,INT_ANY)

		# Execute sql Query
		cursor.execute(query, values)

# Close the cursor
cursor.close()

# Commit the transaction
database.commit()

# Close the database connection
database.close()

# Print results
print ""
print "All Done! Bye, for now."
print ""
