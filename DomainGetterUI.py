import xlsxwriter
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
import os.path
import whois

root = Tk()
root.wm_title("Domain Getter")
root.grid()

singleMultiVar = IntVar()

def singleMultiToggle():
	if (singleMultiVar == 1):
		print("hello")

#singleRadioBtn = Radiobutton(root, variable = singleMultiVar, value = 1, 
#							 command = singleMultiToggle).grid(row = 0, column = 0)

singleDomainLabel = Label(root, text="Single Domain Entry: ").grid(row = 0, column = 0)

singleDomainEntryContent = StringVar()
singleDomainEntry = Entry(root, textvariable = singleDomainEntryContent).grid(row = 0, column = 1)

def getSingleDomainBtnClick():
	domainInput = singleDomainEntryContent.get()
	result = whois.whois(domainInput)
	messagebox.showinfo("Results", result if result.domain_name is not None else "Not Owned \n\nNOTE: simpler domain queries such as \"bing.com\" yield more accurate results than \"http://www.bing.com\"")

getSingleDomainBtn = Button(root, text="Get Info", command=getSingleDomainBtnClick).grid(row = 0, column = 2)




# multi entry through file



#multiRadioBtn = Radiobutton(root, variable = singleMultiVar, value = 2,
#							command = singleMultiToggle).grid(row = 1, column = 0)

multiDomainLabel = Label(root, text = "Multi Domain Entry: ").grid(row = 1, column = 0)

multiDomainEntryContent = StringVar()
multiDomainEntry = Entry(root, textvariable = multiDomainEntryContent).grid(row = 1, column = 1)

def browseFile():
	filePath = filedialog.askopenfilename(filetypes = (("Text files", "*.txt"),
                                   				       ("All files", "*.*") ))
	if (filePath):
		multiDomainEntryContent.set(filePath)

browseFileBtn = Button(root, text = "Browse", command = browseFile).grid(row = 1, column = 2)

def createOutputExcelFile(outputExcelFile, givenDomainNames, domainInfo):
	wb = xlsxwriter.Workbook(outputExcelFile)
	ws = wb.add_worksheet()

	headerFormat = wb.add_format({'bold': True})
	dateFormat = wb.add_format({'num_format': 'mmmm d yyyy'})
	notOwnedFormat = wb.add_format()
	notOwnedFormat.set_pattern(1)
	notOwnedFormat.set_bg_color('red')

	ws.set_column(0, 0, 20)
	ws.set_column(2, 4, 20)
	ws.set_column(5, 6, 60)
	ws.set_column(7, 9, 25)

	headers = ["Input Domain Name", "Owned", "Domain Names", "Registrar", "Name", "Emails", "Address", "Updated Dates", "Creation Date", "Expiration Date"]
	ws.write_row(0, 0, headers, headerFormat)

	for idx, domain in enumerate(givenDomainNames):
		dInfo = domainInfo[idx]

		if (dInfo.domain_name is None):
			ws.write(idx + 1, 0, domain, notOwnedFormat)
			ws.write(idx + 1, 1, "No", notOwnedFormat)
		else:
			ws.write(idx + 1, 0, domain)
			ws.write(idx + 1, 1, "Yes")
			ws.write(idx + 1, 2, str(dInfo.domain_name))
			ws.write(idx + 1, 3, str(dInfo.registrar))
			ws.write(idx + 1, 4, str(dInfo.name))
			ws.write(idx + 1, 5, str(dInfo.emails))
			ws.write(idx + 1, 6, dInfo.address + ", " + dInfo.city + ", " + dInfo.state + " " + dInfo.zipcode if dInfo.address is not None else str(dInfo.address))
			ws.write(idx + 1, 7, str(dInfo.updated_date) if dInfo.updated_date is None else "; ".join([d.strftime("%B %d, %Y") for d in dInfo.updated_date]) if type(dInfo.updated_date) is list else dInfo.updated_date.strftime("%B %d, %Y"))
			ws.write(idx + 1, 8, str(dInfo.creation_date) if dInfo.creation_date is None else "; ".join([d.strftime("%B %d, %Y") for d in dInfo.creation_date]) if type(dInfo.creation_date) is list else dInfo.creation_date.strftime("%B %d, %Y"))
			ws.write(idx + 1, 9, str(dInfo.expiration_date) if dInfo.expiration_date is None else "; ".join([d.strftime("%B %d, %Y") for d in dInfo.expiration_date]) if type(dInfo.expiration_date) is list else dInfo.expiration_date.strftime("%B %d, %Y"))

	wb.close()

multiDomainBtnContent = StringVar()

def getMultiDomainBtnClick():
	multiDomainBtnContent.set("Working...")
	filePath = multiDomainEntryContent.get()
	if (os.path.isfile(filePath)):
		with open(filePath) as f:
			fileContent = [x.strip('\n').strip() for x in f.readlines()]

		domainInfo = [whois.whois(x) for x in fileContent]
		fileDirectory = os.path.dirname(filePath)
		outputExcelFile = fileDirectory + "/DomainInfo"

		if (os.path.isfile(outputExcelFile + ".xlsx")):
			fileNum = 1
			while (os.path.isfile(outputExcelFile + str(fileNum) + ".xlsx")):
				fileNum += 1

			outputExcelFile = outputExcelFile + str(fileNum) + ".xlsx"

		createOutputExcelFile(outputExcelFile, fileContent, domainInfo)
	else:
		messagebox.showinfo("WARNING: Invalid File Path", "The file path provided seems to be invalid, please review and try again.\nIf the file path is correct and you still receive this error, please contact the developer.")

	multiDomainBtnContent.set("Export to Excel")

multiDomainBtnContent.set("Export to Excel")
getMultiDomainBtn = Button(root, text = multiDomainBtnContent.get(),
					       command = getMultiDomainBtnClick).grid(row = 1, column = 3)

root.mainloop()