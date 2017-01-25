import socket
import errno
import xlsxwriter
from Tkinter import *
import tkFileDialog
import tkMessageBox
import os.path
import whois
import time


class DomainGetter(Frame):
	def __init__(self):
		Frame.__init__(self)
		self.master.title("Domain Getter")
		self.master.iconbitmap(os.path.join(os.path.dirname(sys.executable), 'data', 'windowIcon'))

		# initialize grid
		self.master.rowconfigure(2, weight = 1)
		self.master.columnconfigure(4, weight = 1)
		self.grid(sticky = W + E + N + S)

		# initialize grid contents

		# single domain query objects
		self.singleDomainLabel = Label(self, text = "Single Domain Entry: ")
		self.singleDomainLabel.grid(row = 0, column = 0)

		self.singleDomainEntryContent = StringVar()
		self.singleDomainEntry = Entry(self, textvariable = self.singleDomainEntryContent)
		self.singleDomainEntry.grid(row = 0, column = 1)

		self.getSingleDomainBtn = Button(self, text = "Get Info", command = self.getSingleDomainBtnClick)
		self.getSingleDomainBtn.grid(row = 0, column = 2)

		# multi domain query objects
		self.multiDomainLabel = Label(self, text = "Multi Domain Entry: ")
		self.multiDomainLabel.grid(row = 1, column = 0)

		self.multiDomainEntryContent = StringVar()
		self.multiDomainEntry = Entry(self, textvariable = self.multiDomainEntryContent)
		self.multiDomainEntry.grid(row = 1, column = 1)

		self.browseFileBtn = Button(self, text = "Browse", command = self.browseFile)
		self.browseFileBtn.grid(row = 1, column = 2)

		self.getMultiDomainBtn = Button(self, text = "Export to Excel", command = self.getMultiDomainBtnClick)
		self.getMultiDomainBtn.grid(row = 1, column = 3)


	def getSingleDomainBtnClick(self):
		domainInput = self.singleDomainEntryContent.get()
		result = whois.whois(domainInput)
		if (result.domain_name is not None):
			resultContent = ("Domain Names: \n" + str(result.domain_name) + "\n\n"
						  + "Registrar: \n" + str(result.registrar) + "\n\n"
						  + "Name: \n" + str(result.name) + "\n\n"
						  + "Emails: \n" + str(result.emails) + "\n\n"
						  + "Address: \n" + (result.address + ", " + result.city + ", " + result.state + " " + result.zipcode if result.address is not None else str(result.address)) + "\n\n"
						  + "Updated Dates: \n" + (str(result.updated_date) if result.updated_date is None else "; ".join([d.strftime("%B %d, %Y") for d in result.updated_date]) if type(result.updated_date) is list else result.updated_date.strftime("%B %d, %Y")) + "\n\n"
						  + "Creation Date: \n" + (str(result.creation_date) if result.creation_date is None else "; ".join([d.strftime("%B %d, %Y") for d in result.creation_date]) if type(result.creation_date) is list else result.creation_date.strftime("%B %d, %Y")) + "\n\n"
						  + "Expiration Date: \n" + (str(result.expiration_date) if result.expiration_date is None else "; ".join([d.strftime("%B %d, %Y") for d in result.expiration_date]) if type(result.expiration_date) is list else result.expiration_date.strftime("%B %d, %Y")))
			tkMessageBox.showinfo("Results", resultContent)
		else:
			tkMessageBox.showinfo("Results", "Not Owned \n\nNOTE: simpler domain queries such as \"bing.com\" yield more accurate results than \"http://www.bing.com\"")


	def browseFile(self):
		filePath = tkFileDialog.askopenfilename(filetypes = (("Text files", "*.txt"),
	                                   				       ("All files", "*.*") ))
		if (filePath):
			self.multiDomainEntryContent.set(filePath)


	def createOutputExcelFile(self, outputExcelFile, givenDomainNames, domainInfo):
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


	def getMultiDomainBtnClick(self):
		try:
			self.getMultiDomainBtn.config(text = "Working...")
			self.getMultiDomainBtn.update_idletasks()

			filePath = self.multiDomainEntryContent.get()
			if (os.path.isfile(filePath)):
				with open(filePath) as f:
					fileContent = [x.strip('\n').strip() for x in f.readlines()]

				domainInfo = [whois.whois(x) for x in fileContent]
				fileDirectory = os.path.dirname(filePath)
				outputExcelFile = fileDirectory + "/DomainInfo"

				# if the output file name is already taken
				if (os.path.isfile(outputExcelFile + ".xlsx")):
					fileNum = 1
					while (os.path.isfile(outputExcelFile + str(fileNum) + ".xlsx")):
						fileNum += 1
					outputExcelFile = outputExcelFile + str(fileNum) + ".xlsx"
				else:
					outputExcelFile = outputExcelFile + ".xlsx"

				self.createOutputExcelFile(outputExcelFile, fileContent, domainInfo)
			else:
				tkMessageBox.showinfo("WARNING: Invalid File Path", "The file path provided seems to be invalid, please review and try again.\nIf the file path is correct and you still receive this error, please contact the developer.")

		except socket.error as e:
			if e.errno == errno.WSAECONNRESET:
				reconnect()
				retry_action()
			else:
				raise
		except:
			tkMessageBox.showinfo("Error", "Something went wrong. Make sure the path you entered is correct and try again.\nIf the error continues, please contact the developer.")

		self.getMultiDomainBtn.config(text = "Export to Excel")
		if ('outputExcelFile' in locals()):
			tkMessageBox.showinfo("Success", "Results can be found here:\n\n" + outputExcelFile)


if __name__ == "__main__":
	DomainGetter().mainloop()