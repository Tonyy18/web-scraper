import lib
import xlsxwriter
import datetime
import sys
import os

empty_chr = " "

def findBooklists(parentId, periodId):
	lists = lib.getBookLists(parentId, periodId)
	if(lists and len(lists) > 0):
		return lists
	return None

def startScraping():
	organizations = lib.getOrganizations()
	#parentId = ID for previous dropdown value

	results = []
	count = 1
	for org in organizations:

		print()
		print("------------ " + str(count) + " / " + str(len(organizations)) + " ------------")
		print()

		orgName = org["Name"].strip()
		orgId = org["Id"]

		#if(orgName != "Haagse Hogeschool (HHS)"):
			#continue

		final = {
			"name": orgName,
			"departments": []
		}

		print("Organization: " + orgName)
		print()

		#Second dropdown
		departments = lib.getDepartments(orgId)
		for dep in departments:
			depName = dep["Name"]
			depId = dep["Id"]

			depLs = {
				"name": depName,
				"periods": []
			}

			print("Department: " + depName)
			print()

			#Third dropdown (years)
			periods = lib.getPeriods(depId)
			if(periods):
				for period in periods:

					perName = period["Name"]
					perId = period["Id"]

					perLs = {
						"name": perName
					}

					print("Periods: " + perName)

					#Unknown dropdowns
					groups = lib.getGroups(perId)
					if(groups):

						perLs["group1"] = []

						for group in groups:

							#first group (4 fields)
							gName = group["Name"]
							gId = group["Id"]

							g1Ls = {
								"name": gName
							}

							print("Group1: " + gName)

							groups2 = lib.getGroups(perId, {
								"group1": gId
							})

							if(groups2):

								g1Ls["group2"] = []

								for group2 in groups2:
									#second group (5 fields)
									g2Name = group2["Name"]
									g2Id = group2["Id"]

									g2Ls = {
										"name": g2Name
									}

									print("Group2: " + g2Name)

									groups3 = lib.getGroups(perId, {
										"group2": g2Id
									})
									if(groups3):

										for group3 in groups3:
											#Third group (6 field)
											if("group3" not in g2Ls):
												g2Ls["group3"] = []

											g3Name = group3["Name"]
											g3Id = group3["Id"]

											g3Ls = {
												"name": g3Name
											}

											print("Group3: " + g3Name)

											lists = findBooklists(g3Id, perId)
											if(lists):
												for ls in lists:
													print("Book list: " + ls["Name"])
												g3Ls["bookLists"] = lists

											g2Ls["group3"].append(g3Ls)
									else:
										#Booklist
										lists = findBooklists(g2Id, perId)
										if(lists):
											for ls in lists:
												print("Book list: " + ls["Name"])
											g2Ls["bookLists"] = lists

									g1Ls["group2"].append(g2Ls)
							else:
								lists = findBooklists(gId, perId)
								if(lists):
									for ls in lists:
										print("Book list: " + ls["Name"])
									g1Ls["bookLists"] = lists

							perLs["group1"].append(g1Ls)

					depLs["periods"].append(perLs)

				print()

			final["departments"].append(depLs)

		results.append(final)

		writeOrganization(final)
		count = count + 1
	save()

workbook = None
worksheet = None
row = 1
def startWriting():
	global workbook
	global worksheet
	date = datetime.datetime.now()
	name = str(date.day) + "-" + str(date.month) + "-" + str(date.year)
	workbook = xlsxwriter.Workbook("results/" + name + '.xlsx', {'strings_to_urls': False})
	worksheet = workbook.add_worksheet()

def writeRow(lst = [], empty = False):
	global row
	col = "A"
	col_width = {
		"A": 30,
		"B": 15,
		"C": 15,
		"D": 40,
		"E": 50,
		"F": 30,
		"G": 50,
		"H": 50,
		"I": 50,
		"Q": 100
	}
	if(empty == False):
		#Else writes empty row
		for key in lst:
			if(col in col_width):
				worksheet.set_column(col + ":" + col, col_width[col])
			rw = col + str(row)
			worksheet.write(rw, key)
			col = chr(ord(col) + 1)
	row = row + 1

def writeOrganization(org):
	cols = [] #Row to be written
	for a in range(0, 17):
		#Create empty cols in a case value doesn't exist
		cols.append(empty_chr)

	cols[0] = org["name"]
	print("Writing " + org["name"])

	for dep in org["departments"]:
		#second option
		cols[1] = dep["name"]
		if("periods" in dep):
			for per in dep["periods"]:
				#third option
				cols[2] = per["name"]

				if("group1" in per):
					for g1 in per["group1"]:
						#fourth option
						cols[3] = g1["name"]

						if("group2" in g1):
							for g2 in g1["group2"]:
								#fifth option
								cols[4] = g2["name"]

								if("group3" in g2):
									for g3 in g2["group3"]:
										#sixth option
										cols[5] = g3["name"]
										if("bookLists" in g3):
											for bookList in g3["bookLists"]:
												scrapeBookList(cols, bookList)
										else:
											writeRow(cols)
								else:
									#No group3

									if("bookLists" in g2):
										cols[5] = empty_chr
										for bookList in g2["bookLists"]:
											scrapeBookList(cols, bookList)
									else:
										cols[5] = empty_chr
										writeRow(cols)
						else:
							#no group2

							if("bookLists" in g1):
								cols[4] = empty_chr
								cols[5] = empty_chr
								for bookList in g1["bookLists"]:
									scrapeBookList(cols, bookList)
							else:
								cols[4] = empty_chr
								cols[5] = empty_chr
								writeRow(cols)
				else:
					#no group1

					if("bookLists" in per):
						cols[3] = empty_chr
						cols[4] = empty_chr
						cols[5] = empty_chr
						for bookList in per["bookLists"]:
							scrapeBookList(cols, bookList)
					else:
						cols[3] = empty_chr
						cols[4] = empty_chr
						cols[5] = empty_chr
						writeRow(cols)
		else:
			if("bookLists" in dep):
				cols[2] = empty_chr
				cols[3] = empty_chr
				cols[4] = empty_chr
				cols[5] = empty_chr
				for bookList in dep["bookLists"]:
					scrapeBookList(cols, bookList)
			else:
				cols[2] = empty_chr
				cols[3] = empty_chr
				cols[4] = empty_chr
				cols[5] = empty_chr
				writeRow(cols)

def scrapeBookList(cols, bookList):
	url = "https://www.studystore.nl/en/boekenlijst/"
	
	name = cols[0]
	period = cols[2].split("-")[0].strip()
	number = bookList["SchoolListNumber"]

	cols[6] = bookList["Name"]

	url = url + name + "/" + period + "/" + number + "/" + bookList["Url"]

	print("Requesting book list: " + bookList["Url"])

	#scrape books
	html = lib.get(url, soup=True)
	if(html):
		courses = html.find_all("div", {"class": "booklist__course"})
		for course in courses:
				course_name = course.find("h4", {"class": "booklist__course-title"})
				if(course_name):
					cols[7] = course_name.text

				books = course.find_all("div", {"class": "itemlist__row"})
				for book in books:
					book_title = book.find("p", {"class": "booklist__product-title"})
					if(book_title):
						book_name = book_title.text.strip()
						cols[8] = book_name

					warning = book.find("div", {"class": "booklist__product-warning"})
					if(warning):
						ps = warning.find_all("p")
						for p in ps:
							if(p.text.lower() == "compulsory"):
								cols[9] = "Yes"

					authors = book.find("div",{"class": "booklist__product-info"})
					authors = authors.find("div", {"class": "booklist__product-authors"})
					if(authors):
						cols[10] = authors.text.strip()

					subinfo = book.find("div", {"class": "subinfo-sm"})
					if(subinfo):
						subinfo_all = subinfo.findChildren()
						for _subinfo in subinfo_all:
							if("isbn" in _subinfo.text.lower()):
								cols[11] = _subinfo.text.lower().split("isbn:")[1].strip()
							if("edition" in _subinfo.text.lower()):
								cols[12] = _subinfo.text.lower().split("edition:")[1].strip()

					price = ""
					_price = book.find("div", {"class": "product-list__price_from"})
					if(_price):
						price = _price.text
					discount = price
					_discount = book.find("div", {"class": "product-list__price_to"})
					if(_discount):
						discount = _discount.text
					cols[13] = price
					cols[14] = discount

					cols[16] = url
					writeRow(cols)

					cols[8] = empty_chr
					cols[9] = empty_chr
					cols[10] = empty_chr
					cols[11] = empty_chr
					cols[12] = empty_chr
					cols[13] = empty_chr
					cols[14] = empty_chr
					cols[15] = empty_chr
					cols[16] = empty_chr

	cols[6] = empty_chr

def save():
	workbook.close()
	print("Results saved")

if(__name__ == "__main__"):
	try:
		startWriting()
		writeRow([
			"Name of the University",
			"Location",
			"Academic Year",
			"List Name",
			"Programme", "Year/Quarter",
			"School/Institution",
			"Name of the course",
			"Name of publication",
			"Compulsory?",
			"Author",
			"ISBN (13)",
			"Edition",
			"Price (New only - without discount)",
			"Price (New only - after discount)",
			"Alert Text",
			"Link",
		])
		writeRow(empty=True)
		startScraping()
	except KeyboardInterrupt:
		print()
		print('Exiting script ...')
		print()
		save()
		try:
			sys.exit(0)
		except SystemExit:
			os._exit(0)