# Developer: Aseem Sethi, Nov 2021
# Install pip
# $ python3 -m pip install --user --upgrade pip
# $ python3 -m pip --version
# $ python3 -m pip install --user virtualenv
# $ python3 -m venv pat-env
# $ source pat-env/bin/activate
# $ cd pat-env; $ mkdir pat
# $ copy pat.py into this directory
# This is where the PDFs will be kept in their own unique dir for every run
# tmp1 and tmp11 pdf files are temporary files created, that can be deleted later
#
# Install PDFminer and xlrd modules
# pip install pdfminer, xlrd
#
# We need to run ocrmypdf from docker as explained in https://ocrmypdf.readthedocs.io/en/latest/docker.html
# Run the following 2 commands in your terminal where you will run this python exe file
#		$ sudo docker pull  jbarlow83/ocrmypdf
#		$ sudo docker tag jbarlow83/ocrmypdf ocrmypdf
# This is then used as following in the code 
#		"sudo docker run --rm -i ocrmypdf - - <tmp1.pdf >tmp11.pdf"
#
# Run the code
# Ensure that there is a file called resultList1.xls file in the "pat" directory.
# This same xls file will be modified by addition of 3 columns at the end of the run.
# $ python pat.py
#
import xlrd
from pathlib import Path
from datetime import datetime
import shutil
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
		if i > 12:
			break;
	
	# Get Agent Name and Address
	agent = ""
	start = text.lower().find("agent") 
	#print (start)
	i=1
	for line in text[start+7:].split("\n"):
		#print(line)
		agent += line + ", "
		i=i+1
		if i > 20:
			break;

	# Freeing Up
	device.close()
	sio.close()
	return email, address, agent

def workon(sh, rowx, timeNow):
	PDFDir = "PDF" + timeNow.strftime("%d-%m-%Y-%H:%M:%S")
	Path(PDFDir).mkdir(parents=True, exist_ok=True)
	print("############ Processing row.. {0}".format(rowx))
	print("URL: {0}".format(sh.cell_value(rowx, colx=0)))
	link = sh.hyperlink_map.get((rowx, 0))
	# Url is of type http://patentscope.wipo.int/search/en/WO2021208467
	url = '(No URL)' if link is None else link.url_or_path
	print("URL: {0}".format(url))
	if url == '(No URL)':
		print ("..exiting, no URL found in 1st Column")
		return "Start URL null", "Start URL null", "Start URL null"
	pdfName = url.split('/')[-1]  # get the document id
	# print (pdfName)
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
	if start == -1:
		# We did not find the RO101, lets retry
		print ("...no RO/101 found using tab=PCTDOCUMENTS url")
		urlDoc = 'https://patentscope.wipo.int/search/en/detail.jsf?docId=' + url.split('/')[-1]
		#urlDoc += '#detailMainForm:MyTabViewId:PCTDOCUMENTS'
		print (".....trying: " + urlDoc)
		data1 = urllib.parse.urlencode(
		{
		"javax.faces.partial.ajax": "true",
		"javax.faces.source": "detailMainForm:MyTabViewId",
		"javax.faces.partial.execute": "detailMainForm:MyTabViewId",
		"javax.faces.partial.render": "detailMainForm:MyTabViewId",
		"javax.faces.behavior.event": "tabChange",
		"javax.faces.partial.event": "tabChange",
		"detailMainForm:MyTabViewId_contentLoad": "true",
		"detailMainForm:MyTabViewId_newTab":"detailMainForm:MyTabViewId:PCTDOCUMENTS",
		"detailMainForm:MyTabViewId_tabindex": "6",
		"detailMainForm": "detailMainForm",
		"detailMainForm:MyTabViewId_activeIndex": "6"
		})
		data1 = data1.encode('ascii')
		req=urllib.request.Request(urlDoc)
		req.add_header("X-Requested-With", "XMLHttpRequest")
		contents = urllib.request.urlopen(req, data=data1).read()
		contents1 = contents.decode("utf-8") 
		#print (contents1)
		#file = open("tmp99.txt", 'w')
		#file.write(contents1)
		#file.close()
		start = contents1.find("(RO/101)") 
		if start == -1:
			print ("...no RO/101 found again")
			return "RO101 not found", "RO101 not found", "RO101 not found"
	#end = contents1.find("\">PDF ", start)
	start = contents1.find("a href=\"", start)
	end = contents1.find("\" class=", start)
	#matches=re.findall(r'\"(.+?)\"',text)
	#print(start, end)
	#substring = contents1[start+73:end]
	substring = contents1[start+8:end]
	urlRO101 = "https://patentscope.wipo.int/" + substring
	print("Getting RO/101 PDF: " + urlRO101)
	try:
		contentsRO101 = urllib.request.urlopen(urlRO101).read()
	except:
		print("Invalid RO101 URL")
		return "RO101  error", "RO101 error", "RO101 error"
	file = open("tmp1.pdf", 'wb')
	file.write(contentsRO101)
	file.close()
	newFile = PDFDir +'/' + pdfName + '.pdf'
	shutil.copy2('tmp1.pdf', newFile)
	print("Running docker cmd to OCR read the pdf file..")
	os.system('sudo docker run --rm -i ocrmypdf - - <tmp1.pdf >tmp11.pdf')
	email, address, agent = pdf_from_file_to_txt("tmp11.pdf")
	print ("Email: " + email + ", Address: " + address + ", Agent: " + agent)
	return email, address, agent


file=u'resultList1.xls'
now = datetime.now()
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
	email, address, agent = workon(sh, rx, now)
	w_sheet.write(rx, sh.ncols, email)
	w_sheet.write(rx, sh.ncols+1, address)
	w_sheet.write(rx, sh.ncols+2, agent)
	wb.save('resultList1.xls')