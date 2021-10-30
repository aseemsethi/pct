# Install and Activate venv
# pip install pdfminer, xlrd
# We need to run ocrmypdf from docker as explained in https://ocrmypdf.readthedocs.io/en/latest/docker.html
# Run the following 2 commands in your terminal where you will run this exe file
#		$ sudo docker pull  jbarlow83/ocrmypdf
#		$ sudo docker tag jbarlow83/ocrmypdf ocrmypdf
# This is then used as following in the code 
#		"sudo docker run --rm -i ocrmypdf - - <tmp1.pdf >tmp11.pdf"

# Install 
# tar -zxvf leptonica-1.82.0.tar.gz 
# cd leptonica-1.82.0/,    ./configure, make, sudo make install sudo ldconfig
# export LD_LIBRARY_PATH=/home/ec2-user/environment/pat-venv/lept/leptonica-1.82.0/src/.libs

import xlrd
import urllib.request
import re
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfpage import PDFPage
from io import StringIO, BytesIO
import os
from xlutils.copy import copy


def pdf_from_file_to_txt(fileName):
	# PDFMiner Analyzers
	rsrcmgr = PDFResourceManager()
	sio = StringIO()
	codec = "utf-8"
	laparams = LAParams()
	device = TextConverter(rsrcmgr, sio, laparams=laparams)
	interpreter = PDFPageInterpreter(rsrcmgr, device)
	 
	# path to our input file
	pdf_file = fileName
	 
	# Extract text
	pdfFile = open(pdf_file, "rb")
	for page in PDFPage.get_pages(pdfFile, 
		caching=True, check_extractable=True):
	   interpreter.process_page(page)
	pdfFile.close()
	 
	# Return text from StringIO
	text = sio.getvalue()
	#print(text)
	email = ""
	for item in text.split("\n"):
		if "@" in item:
			email += item.strip().replace(" ", "")
			email += ", "
			print (email)
	
	# Get "Applicant Only" - Address
	address = ""
	start = text.find("States") 
	#print (start)
	i=1
	for line in text[start+10:].split("\n"):
		#print(line)
		address += line + ", "
		i=i+1
		if i > 10:
			break;

	# Freeing Up
	device.close()
	sio.close()
	return email, address

def workon(sh, rowx):
	print("############ Processing row.. {0}".format(rowx))
	print("URL: {0}".format(sh.cell_value(rowx, colx=0)))
	link = sh.hyperlink_map.get((rowx, 0))
	# Url is of type http://patentscope.wipo.int/search/en/WO2021208467
	url = '(No URL)' if link is None else link.url_or_path
	print("URL: {0}".format(url))
	# print(url.split('/')[-1])  # get the document id
	# Convert url to 
	# #https://patentscope.wipo.int/search/en/detail.jsf?docId=WO2021208467&tab=PCTDOCUMENTS
	urlDoc = \
		'https://patentscope.wipo.int/search/en/detail.jsf?docId=' + \
		url.split('/')[-1] + \
		'&tab=PCTDOCUMENTS'
	print("Getting PCTDOCUMENTS: " + urlDoc)
	contents = urllib.request.urlopen(urlDoc).read()
	contents1 = contents.decode("utf-8") 
	start = contents1.find("(RO/101)") 
	end = contents1.find("\">PDF ", start)
	#matches=re.findall(r'\"(.+?)\"',text)
	# print(start, end)
	substring = contents1[start+73:end]
	urlRO101 = "https://patentscope.wipo.int/" + substring
	print("Getting RO/101 PDF: " + urlRO101)
	contentsRO101 = urllib.request.urlopen(urlRO101).read()
	file = open("tmp1.pdf", 'wb')
	file.write(contentsRO101)
	file.close()
	print("Running docker cmd to OCR read the pdf file..")
	os.system('sudo docker run --rm -i ocrmypdf - - <tmp1.pdf >tmp11.pdf')
	email, address = pdf_from_file_to_txt("tmp11.pdf")
	print ("Email: " + email + ", Address: " + address)
	return email, address


file=u'resultList1.xls'
try:
    book = xlrd.open_workbook(file,
    	encoding_override="cp1251", formatting_info=True)  
except:
    book = xlrd.open_workbook(file)
print("The number of worksheets is {0}".format(book.nsheets))
print("Worksheet name(s): {0}".format(book.sheet_names()))
sh = book.sheet_by_index(0)
#start processing
print("Sheet: {0} Rows:{1} Cols:{2}".format(sh.name, sh.nrows, sh.ncols))
wb = copy(book) # a writable copy (I can't read values out of this, only write to it)
w_sheet = wb.get_sheet(0) # the sheet to write to within the writable copy
for rx in range(sh.nrows):
	if rx < 3:
		continue
	#workon(sh.row(rx))
	email, address = workon(sh, rx)
	w_sheet.write(rx, sh.ncols, email)
	w_sheet.write(rx, sh.ncols+1, address)
	wb.save('resultList1.xls')
	