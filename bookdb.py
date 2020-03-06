##Function which will use pip3 to install any missing packages on import
import logging
logging.basicConfig(filename='bookdb.log',format='%(asctime)s  %(levelname)-8s [%(threadName)s-%(name)s] %(message)s',level=logging.DEBUG)
logger = logging.getLogger("Main")
def install_and_import(package):
    import importlib
    try:
        importlib.import_module(package)
    except ImportError:
        logger.info("Installing %s",package)
        import pip
        subprocess.call(['pip3', 'install', package])
    finally:
        globals()[package] = importlib.import_module(package)
import subprocess
import time
import win32com.client
##lets us manipulate excel spreadsheets w/o excel
install_and_import('openpyxl')
##used to ssh into server that has calibre
install_and_import('paramiko')
##if excel is installed locally (and a windows machine) can open excel to do sorts
install_and_import('win32com.client')
##upload spreadsheet to dropbox and download from dropbox
install_and_import('dropbox')	
import queue
import workbook
import calibre
import sys
from threading import Thread
import concurrent.futures as future

##writes message for each blank line
def log_blanks(result,previous):
	blank = False
	for line in result:
		##If any blank lines then open the blanks file
		if line[0] == "" or line[1] == "":
			blank = True
	if blank:
		try:
			##Open missing.txt in append mode, create it if it doesn't exist
			blanks = open("missing.txt", "a+")
			for i,line in enumerate(result):
				if line[0] == "" or line[1] == "":
					##Depending on information, write the file with the relevant info
					logger.warning("Adding note about missing information on %s",line[3])
					logger.debug("Full information is %s",line)
					if i == 0:
						if previous == None:
							blanks.write("First item scanned, "+line[3]+" had missing data\n")
						else:
							blanks.write(line[3]+" which was scanned after "+previous[3]+", written by "+previous[0]+" had missing data\n")
					else:
						blanks.write(line[3]+" which was scanned after "+result[i-1][3]+", written by "+result[i-1][0]+" had missing data\n")
		finally:
			blanks.close()
			return True
	return False
	
##Thread Manager monitors input queue and spawns sub threads to look up metadata. Auto saves every 10 inputs
def manageThread(q, wb, file, books, client):
	logger = logging.getLogger("ThreadManager")
	blank = False
	##Used to track running sub threads
	running = []
	result = []
	previous = None
	while(True):
		logger.info("Thread Manager is listening")
		##Blocking call for the next input
		temp = q.get()
		logger.debug("Received input %s",temp)
		##If the input is numeric assume it's an isbn and do a search
		if temp.isnumeric():
			##spawn a new thread to add book to database
			running.append(future.ThreadPoolExecutor(thread_name_prefix='lookup '+temp).submit(workbook.addBook,client,temp,books))
		##if the input is the word exit then close all threads and exit
		elif temp == "exit":
			logger.debug("Waiting on metadata to finish")
			##wait on all threads in running to finish
			for i,x in enumerate(running):
				while len(result)-1 < i:
					result.append(None)
				result[i] = x.result()
			logger.debug("Checking for blanks before exiting")
			if log_blanks(result,previous): blank = True
			logger.debug("Thread Manager is Exiting")
			return blank
		##After every new input check if the length of running is > 9 if it is, save workbook before adding new threads
		if len(running)>9:
			logger.debug("Background Saving")
			for i,x in enumerate(running):
				while len(result)-1 < i:
					result.append(None)
				result[i] = x.result()
			logger.debug("Checking for blanks before continuing")
			if log_blanks(result,previous): blank = True
			previous = result[len(result)-1]
			workbook.save(wb, file)
			running = []
			result = []
	##Sanity check, manage Thread should never reach here
	logger.critical("manageThread left while loop")
	print("Critical Error, manager has failed\a")
	exit()

##Always assume we will be connecting to a server
local = False
logger.info("----------------------------------------------------------")
try:
	##Look for configuration File
	settings = open("bookdb.conf", "r")
	##First Line is the IP address of the server running calibre (optional)
	ip = settings.readline().strip()
	logger.info("ip address is %s",ip)
	print("Connecting to "+ip)
	##Username to SSH into server (optional)
	username = settings.readline().strip()
	logger.info("username is %s",username)
	print("As "+username)
	##Key to SSH into server (optional)
	key = settings.readline().strip()
	logger.info("key location is %s",key)
	print("using key at "+key)
	##Location to store database (required)
	file = settings.readline().strip()
	logger.info("Workbook is %s",file)
	print("using database at "+file)
	##Logging Level (required)
	level = settings.readline().strip()
	print("logger at",level)
	logger.info("Setting log level to %s",level)
	logging.getLogger().setLevel(level)
	##Dropbox Token (optional)
	token = settings.readline().strip()
	if token:
		print("Dropbox Token Found")
		logger.info("Found Access Token")
	if not ip or not username or not key or not file or not level:
		##If any fields are blank, go to exception
		raise OSError
except Exception:
	if file and level:
		##If required information is present, fallback to local mode
		print("Invalid information for remote host, attempting to use localhost")
		logger.exception("Invalid Remote Host information.  Attempting to use localhost")
		local = True
	else:
		##If required information is absent, request all information and write it to settings file
		settings.close()
		print("No Valid Settings Found")
		logger.warning("No Valid Settings Found")
		settings = open("bookdb.conf", "w")
		ip = input("Enter the servers IP address: ")
		username = input("Enter the username: ")
		key = input("Enter the path to the ssh key: ")
		file = input("Enter the excel workbook: ")
		level = input("Enter the log level: ")
		logger.info("Writing %s %s %s %s %s to bookdb.conf",ip,username,key,file,level)
		settings.write(ip+'\n')
		settings.write(username+'\n')
		settings.write(key+'\n')
		settings.write(file+'\n')
		settings.write(level+'\n')
		logger.getLogger().setLevel(level)
	settings.close()
if not local:
	try:
		##If in remote mode, SSH into server
		logger.debug("Attempting to SSH into %s",ip)
		print("Connecting to",ip)
		client = paramiko.SSHClient()
		client.load_system_host_keys()
		client.set_missing_host_key_policy(paramiko.WarningPolicy)
		client.connect(ip, port=22, username=username, password="", pkey=None, key_filename=key)
	except Exception as inst:
		##If SSH fails for any reason fall back to local mode
		print("Unable to connect to host.")
		logger.critical("Failed to connect to %s using %s and %s error was:\n%s",ip,username,key,inst)
		print(inst)
		logger.critical("Falling back to local mode")
		print("falling back to local mode\a")
		client = None
else:
	##Set SSH client to none for local mode
	logger.debug("Skipping connecting to remote host")
	client = None
try:
	if token:
		##If Dropbox token is present, ask if we should download the spreadsheet or use a local one
		question = input("Dropbox token found, Download file?[yes]: ")
		if question.lower().strip() == "yes" or question.lower().strip() == "":
			logger.info("Attempting to download from dropbox")
			workbook.downloadFromDropbox(file, token)
	##Whether or not it was downloaded from dropbox, open the file locally and the Books Tab
	logger.info("Attempting to open file")
	wb = workbook.open_workbook(file)
	books = wb['Books']
	print("Opened",file)
	logger.info("Opened %s",file)
except Exception as inst: 
	##If we can't open the spreadsheet or the books file, creating a new spreadsheet and making the books tab
	print("Unable to open file",file,"error:",inst,"making new file")
	logger.warning("Can't open file %s error: %s, making new file",file,inst)
	wb = workbook.open_workbook(None)
	books = wb.active
	books.title = "Books"
	##Add header row
	row = ("Author","Title","Series","ISBN","Location","Condition","Borrowed By","Date Borrowed")
	books.append(row)
	workbook.save(wb,file)
##Create the blocking queue and spawn the Thread Manager Thread
logger.info("Starting Thread Manager")
q = queue.Queue()
manager = future.ThreadPoolExecutor(thread_name_prefix='TM').submit(manageThread,q,wb,file,books,client)
##Loop forever until exit or quit is typed in
while(True):
	isbn = input("Enter ISBN: ")
	if isbn == "exit" or isbn == "quit":
		##Once told to exit, tell Thread Manager to wrap it up, then save and close the workbook before exiting
		print("waiting for metadata to finish downloading")
		logger.info("Waiting for metadata to finish downloading")
		q.put("exit")
		if manager.result(): any_blanks = True
		else: any_blanks = False
		print("Saving Excel File")
		workbook.saveAndClose(wb, books, file, token)
		if any_blanks:
			print("There was missing data.  Results have been saved to missing.txt")
			try:
				subprocess.Popen([r'C:\Program Files (x86)\Notepad++\notepad++.exe', 'missing.txt'])
			except Exception as inst:
				print("Tried to open missing.txt using Notepad++ got error:",inst)
				logger.exception("Failed to open missing.txt error was %s",inst)
			try:
				excel = win32com.client.Dispatch("Excel.Application")
				logger.debug("Attempting to open %s",file)
				##Load file and Worksheet
				wb = excel.Workbooks.Open(Filename=file)
				ws = wb.Worksheets('Books')
				##Grab Table
				tab = ws.ListObjects("MyBooks")
				##Remove current sorts, and sort by Author,Series,Title
				tab.Sort.SortFields.Clear()
				tab.Sort.SortFields.Add(Key=ws.Range("MyBooks[Author]"), Order=1)
				tab.Sort.SortFields.Add(Key=ws.Range("MyBooks[Series]"), Order=1)
				tab.Sort.SortFields.Add(Key=ws.Range("MyBooks[Title]"), Order=1)
				##Apply Sort, Save, Then close Excel
				tab.Sort.Apply()
			except Exception as inst:
				print("Tried to open",file,"got error:",inst)
				logger.exception("Failed to open %s error was %s",file,inst)
			excel.Visible = True
			while True:
				try:
					win32com.client.GetActiveObject("Excel.Application")
					time.sleep(5)
				except:
					break;
			try:
				wb.Save()
			except Exception as inst:
				wb = excel.Workbooks.Open(Filename=file)
				ws = wb.Worksheets('Books')
			finally:
				##Grab Table
				tab = ws.ListObjects("MyBooks")
				##Remove current sorts, and sort by Author,Series,Title
				tab.Sort.SortFields.Clear()
				tab.Sort.SortFields.Add(Key=ws.Range("MyBooks[Author]"), Order=1)
				tab.Sort.SortFields.Add(Key=ws.Range("MyBooks[Series]"), Order=1)
				tab.Sort.SortFields.Add(Key=ws.Range("MyBooks[Title]"), Order=1)
				tab.Sort.Apply()
				wb.Save()
				excel.Application.Quit()
				workbook.uploadToDropbox(file,token)
			
		client.close()
		logger.info("Exiting Program")
		exit()
	if isbn.isnumeric():
		##If it's numeric put it in the queue for threadManager to use
		logger.debug("Searching for %s",isbn)
		q.put(isbn)
	else:
		##Otherwise give an error and wait for more input
		logger.warning("Error %s is not a valid isbn",isbn)
		print("ISBN's are only numeric\a")