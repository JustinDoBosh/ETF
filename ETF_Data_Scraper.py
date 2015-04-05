import urllib2
from urllib2 import urlopen
from bs4 import BeautifulSoup
from Tkinter import *
import xlsxwriter
import os

master = Tk()
#Need to open the excel file first, because if we open it inside the allETFInfo class it only writes the last etf entered
workbook = xlsxwriter.Workbook('ETFRatings.xlsx')
worksheet = workbook.add_worksheet()

#------------------------------------ GUI ------------------------------------------------------------------------------------------
class GUI:

	def __init__(self):
		self.etfsGUIInput = StringVar()
		self.rootURLNum = IntVar()
		self.rootURLStr = StringVar()
		self.GUIETFList = []
		self.master = master
		self.master.title("ETF Data Scraper")

	def init_window(self):
		directions = StringVar()
		label = Label(master, textvariable=directions )
		directions.set("Select the site you want to scrape, and enter up to 25 ETFs at a time.\n Seperate each ETFs with a comma.\n")
		label.pack()

		etfEntry = Entry(master, textvariable=self.etfsGUIInput)
   			

		R1 = Radiobutton(master, text="etf.com", variable=self.rootURLNum, value=1)
		R1.pack( padx=5, pady=5)

		R2 = Radiobutton(master, text="maxfunds.com", variable=self.rootURLNum, value=2)
		R2.pack( padx=5, pady=5)

		R3 = Radiobutton(master, text="Smartmoney.com", variable=self.rootURLNum, value=3)
		R3.pack( padx=5, pady=10)


		def cleanAndReturnListofEtfs():
			self.GUIETFList = self.etfsGUIInput.get()
			self.GUIETFList = self.GUIETFList.split(', ')
			#print self.GUIETFList
			#print type(self.GUIETFList)
			self.rootURLNum = self.rootURLNum.get()
			if(self.rootURLNum == 1):
				self.rootURLStr = "http://www.etf.com/"
			elif(self.rootURLNum == 2):
				self.rootURLStr = "http://www.maxfunds.com/funds/data.php?ticker="
			elif(self.rootURLNum == 3):
				self.rootURLStr = "http://www.marketwatch.com/investing/Fund/"
			

			baseURL = self.rootURLStr
			etfList =  self.GUIETFList
			row = 0

			for etfSymbol in etfList:
				row += 1
				print "-------Starting Data Collection for " + etfSymbol + " ---------"
				myEtf = getEtfInfo(etfSymbol, row, baseURL)
				myEtf.getData()
				#use an if statement to find out which website we are scraping
 				if(baseURL == "http://www.etf.com/"):
 					myEtf.etfDotComInfo()

 				elif(baseURL == "http://www.maxfunds.com/funds/data.php?ticker="):
 					myEtf.maxfundsDotComInfo()

 				elif(baseURL == "http://www.marketwatch.com/investing/Fund/"):
					myEtf.smartmoneyDotComeInfo()
				print "-------Data Collection Complete for " + etfSymbol + " ---------"

			#close the window 
			master.destroy()

		etfSubmitBtn = Button(master, text="Get Data", command=cleanAndReturnListofEtfs)

		etfEntry.pack(padx=5, pady=5)
		etfSubmitBtn.pack(padx=5, pady=5)	
#------------------------------------ getEtfInfo ------------------------------------------------------------------------------------------
class getEtfInfo:
	def __init__(self, etfSymbol, row, baseURL):
		self.etfSymbol = etfSymbol
		self.row = row 
		self.baseURL = baseURL

	def getData(self):
		#****The 3 web URLs aviable to scrape*****
		#maxfunds: http://www.maxfunds.com/funds/data.php?ticker=VTSMX
		#etf.com: http://www.etf.com/spy
		#Smartmoney: http://www.marketwatch.com/investing/Fund/OARMX

		#get document source code 
		website = urllib2.urlopen(self.baseURL + self.etfSymbol)
		sourceCode = website.read()
		#make soup a global var, so it can be accessed later
		global soup
		soup = BeautifulSoup(sourceCode)

	def etfDotComInfo(self):
		#Test funds: spy, qqq, vti, ivv, GLD, VOO, EEM
		# Widen the first column to make the text clearer.
		worksheet.set_column('A:E', 30)
		#Add formating
		format = workbook.add_format()
		format.set_text_wrap()
		format.set_font_size(14)
		format.set_font_name('Arial')
		format.set_align('center')
		# Write some data headers.
 		worksheet.write('A1', 'ETF Name', format)
 		worksheet.write('B1', 'Time Stamp', format)
 		worksheet.write('C1', 'Efficiency', format)
 		worksheet.write('D1', 'Tradability', format)
 		worksheet.write('E1', 'Fit', format)
 		# Start from the first cell below the headers.
 		row = self.row
 		col = 0

		#parse document to find etf name 
		etfName = soup.find('h1', class_="etf")
		#extract etfName contents (etfTicker & etfLongName)
		etfTicker = etfName.contents[0]
		etfLongName = etfName.contents[1]
		etfTicker = str(etfTicker)
		etfLongName = etfLongName.text
		etfLongName = str(etfLongName)
		etfFullName = etfTicker + ' - ' + etfLongName
		etfFullName = str(etfFullName)
		#print etfFullName

		#get the time stamp for the data scraped 
		etfInfoTimeStamp = soup.find('div', class_="footNote")
		dataTimeStamp = etfInfoTimeStamp.contents[1]
		formatedTimeStamp =  'As of ' + dataTimeStamp.text
		formatedTimeStamp = str(formatedTimeStamp)
		#print formatedTimeStamp

		#create vars 
		etfScores = []
		cleanEtfScoreList = []
		#parse document to find all divs with the class score
		etfScores = soup.find_all('div', class_="score")
		#loop through etfScores to clean them and add them to the cleanedEtfScoreList
		for etfScore in etfScores:
			strippedEtfScore = etfScore.string.extract()
			strippedEtfScore = str(strippedEtfScore)
			cleanEtfScoreList.append(strippedEtfScore)
		#turn cleanedEtfScoreList into a dictionary for easier access
		
		ETFInfoToWrite = [etfFullName, formatedTimeStamp, int(cleanEtfScoreList[0]), int(cleanEtfScoreList[1]), int(cleanEtfScoreList[2])]

		for etf in ETFInfoToWrite:
			worksheet.write(self.row, col, etf, format)
			col += 1
		col = 0

	def maxfundsDotComInfo(self):
		#Test funds: VTIAX, PTTRX, PRFDX, DBLTX, TGBAX, FCNTX
		# Widen the first column to make the text clearer.
		worksheet.set_column('A:B', 40)
		#Add formating
		format = workbook.add_format()
		format.set_text_wrap()
		format.set_font_size(14)
		format.set_font_name('Arial')
		format.set_align('center')
		# Write some data headers.
 		worksheet.write('A1', 'ETF Name', format)
 		worksheet.write('B1', 'Max Rating', format)
 		# Start from the first cell below the headers.
 		row = self.row
 		col = 0
 		#get ETFs name
 		etfName = soup.find('div', class_="dataTop")
 		etfName = soup.find('h2')
 		etfName = str(etfName.text)
 		#get ETFs Max rating score
 		etfMaxRating = soup.find('span', class_="maxrating")
 		etfMaxRating = str(etfMaxRating.text)

 		#create array to store name and rating 
 		ETFInfoToWrite = [etfName, int(etfMaxRating)]

 		for etf in ETFInfoToWrite:
			worksheet.write(self.row, col, etf, format)
			col += 1
		col = 0

	def smartmoneyDotComeInfo(self):
		#Test funds: OAKLX, OAKGX, OARMX, OAKBX, OAKIX, OARIX
		# Widen the first column to make the text clearer.
		worksheet.set_column('A:G', 30)
		#Add formating
		format = workbook.add_format()
		format.set_text_wrap()
		format.set_font_size(14)
		format.set_font_name('Arial')
		format.set_align('center')
		# Write some data headers.
		worksheet.write('A1', 'Fund Name', format)
 		worksheet.write('B1', 'Ticker Symbol', format)
 		worksheet.write('C1', 'Total Return', format)
 		worksheet.write('D1', 'Consistent Return', format)
 		worksheet.write('E1', 'Preservation', format)
 		worksheet.write('F1', 'Tax Efficiency', format)
 		worksheet.write('G1', 'Expense', format)
 		# Start from the first cell below the headers.
 		row = self.row
 		col = 0
 		#get etf Name
 		etfName = soup.find('h1', id="instrumentname")
 		etfName = str(etfName.text)
 		#get etf Ticker
 		etfTicker = soup.find('p', id="instrumentticker")
 		etfTicker = str(etfTicker.text)
 		etfTicker = etfTicker.strip()

 		cleanedLipperScoreList = []
 		cleanedLipperScoreList.append(etfName)
 		cleanedLipperScoreList.append(etfTicker)

 		#get Lipper scores ***NEEDS REFACTORING***
 		lipperScores = soup.find('div', 'lipperleader')
 		lipperScores = str(lipperScores)
 		lipperScores = lipperScores.split('/>')
 		for lipperScore in lipperScores:
 			startIndex = lipperScore.find('alt="')
 			startIndex = int(startIndex)
 			endIndex = lipperScore.find('src="')
 			endIndex = int(endIndex)
 			lipperScore = lipperScore[startIndex:endIndex]
 			startIndex2 = lipperScore.find('="')
 			startIndex2 = startIndex2 + 2
 			endIndex2 = lipperScore.find('" ')
 			#At this point I have the ex: "Total Return: 5"
 			lipperScore = lipperScore[startIndex2:endIndex2]
 			seperatorIndex = lipperScore.find(':')
 			endIndex3 = seperatorIndex
 			startIndex3 = seperatorIndex + 1

 			lipperScoreNumber = lipperScore[startIndex3:]
 			if lipperScoreNumber == '' and lipperScoreNumber == '':
 				pass
 			else:
 				cleanedLipperScoreList.append(int(lipperScoreNumber))

 		for cleanedLipperScore in cleanedLipperScoreList:
			worksheet.write(self.row, col, cleanedLipperScore, format)
			col += 1
		col = 0
#------------------------------------ CallToGo ------------------------------------------------------------------------------------------
#Starts the application 
def callToGo():
	#Sets the height and width of the window
	master.geometry("600x300") 
	#Inits the application 
	app = GUI()
	#sets up the Tkinter window
	app.init_window()
	#Starts Tkinter
	master.mainloop()
	#close the workbook after the all the data is pulled and written to the excel file
	workbook.close()
	#opens the excel file (tested on mac, but not on windows)
	os.system("open ETFRatings.xlsx")

#Starts the application 
callToGo()

