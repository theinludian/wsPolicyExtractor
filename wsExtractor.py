from xlrd import open_workbook,XL_CELL_TEXT
from os.path import join, dirname, abspath
import re,csv,os,sys,argparse

parser = argparse.ArgumentParser(description='Websense Category Parser')

parser.add_argument('-f', action='store', dest='sheetFile', help='Enter the spreadsheet filename.')
parser.add_argument('-s', action='store', dest='catName', help='Enter the category name to parse. NOTE: If space is present use quotes to encapsulate the category name')
parser.add_argument('-o', action='store', dest='csvFile', help='Enter the csv filename.')

results = parser.parse_args()

# Opens the workbook
wb = open_workbook(results.sheetFile)
sheet = wb.sheet_by_index(0)

# Gets the row index of Protocol Filters heading
def getCategoryPosition(catName):
	for idx, cell in enumerate(sheet.col_slice(0,0)):
		if catName in str(cell.value):
			return idx


# ---getFirstProtocol---
# Input:
# -Start position to start searching from
# -sItem :: search Item - will look for the heading with that name
#
# Output:
# Returns the row index of the found item.
def getFirstProtocol(startPos, sItem):
	for idx, cell in enumerate(sheet.col_slice(1,startPos)):
		if sItem in str(cell.value):
			return startPos + idx


# getColumnHeaders:
# Input:
# -Takes one paramter as input - the row number where the heading "Protocol Filters" is located at
#
# Output:
# -Returns a list of all the headers for the csv file
# -The row number of the first double heading -> used later to calculate the blocks of start and end of headings
# -Returns the first legit column heading row number
# -Returns the last column heading value
def getColumnHeaders(startPos):
	headerList = dict({})

	ignoreCells = ["Protocol Actions", "Category Actions", "Usage Count", '', "Time Periods", "Client Count"]
	firstDoublePos = ""
	lastDouble = ""
	initialFirstPosBol = True
	firstHeading = True
	count = 0
	firstHeadingPos = 0
	name = ""

	for idx, cell in enumerate(sheet.col_slice(1,startPos,startPos+25)):
		if cell.value in headerList.values() and initialFirstPosBol == True:
			firstDoublePos = idx + startPos
			lastDoublePos = headerList.get(count-1)
			initialFirstPosBol = False

		elif cell.value in headerList.values():
			pass

		elif ignoreCells[0] in cell.value:
			pass

		elif ignoreCells[1] in cell.value:
			pass

		elif ignoreCells[2] in cell.value:
			pass

		elif ignoreCells[3] == cell.value:
			pass

		elif ignoreCells[4] in cell.value:
			pass

		elif ignoreCells[5] == cell.value:
			pass

		else:
			headerList[count] = cell.value
			count += 1
			if firstHeading == True:
				firstHeadingPos = idx + startPos

	return headerList

# ---getEndPos---
# Input:
# -Give start position from where to start looking for the end of the Category.
#
# Output:
# -Returns the position of the end of the current Category.
def getEndPos(startPos):
	for idx, cell in enumerate(sheet.col_slice(0,startPos)):
		if re.search('[a-zA-Z]+',str(cell.value)) and cell.value != '':
			return startPos + idx


# ---getData---
# Input:
# -The start position from where the Category heading is
# -The position of where the Category ends.
# -The list of headers
#
# Output:
# -Returns a dictionary {rowNumber:csvRow}
def getData(startPos,endCatPos,headerList):
	itemCount = 0
	blockCount = 0
	dataDict = dict({})
	size = (len(headerList.keys()) - 1)
	name = ""
	row = ""

	for idx, cell in enumerate(sheet.col_slice(1,startPos,endCatPos)):

		for item in headerList.values():
		    if item == cell.value:
				if "Name" == cell.value:
					name = sheet.cell(startPos + idx,2).value
					row = name
					blockCount = 1

				elif blockCount == size:
					row += ";" + str(sheet.cell(startPos + idx, 2).value)
					dataDict[itemCount] = row
					itemCount += 1
					row = name
					blockCount = 1


				else :
					row += ";" + str(sheet.cell(startPos + idx, 2).value)
					blockCount += 1

		else:
			pass

	return (dataDict)

# ---toCSV---
# Writes the csv file
#
# Input:
# -the csv headers
# -the csv row dictionary
def toCSV(headers,dataDict):
	f = open(results.csvFile, 'w')
	writer = csv.writer(f, delimiter=';')
	writer.writerow(headers.values())

	for idx,item in enumerate(dataDict):
		f.write(dataDict.get(idx).encode('utf8') + "\n")

	f.close()

# The natural order of the script:
# 1. Get the start position of the specified category
# 2. Get the end position of the specified category
# 3. Get the colum headers
# 4. Populate the dictionary of csv rows
# 5. Write the dictionary to the file
startPos = getCategoryPosition(results.catName)
endPos = getEndPos(startPos + 1)
colHeaderResults = getColumnHeaders(startPos)
dataResults = getData(startPos, endPos, colHeaderResults)
toCSV(colHeaderResults,dataResults)

print
print "Printed to csv - ", results.csvFile
print
