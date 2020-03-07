# Function which will use pip3 to install any missing packages on import
import logging
import my_logger
import subprocess
import time
import win32com.client
import traceback
import queue
from os import path
import workbook
from install_and_import import install_and_import
import concurrent.futures as future
import paramiko
import win32com.client

# used to ssh into server
install_and_import('paramiko')
# if excel is installed locally (and a windows machine) can open excel to do sorts
install_and_import('win32com.client')
file_path = ""


# Checks if the settings are valid
def check_settings(file_path_in, local_in, ip_in, username_in, key_in, file_in, level_in, token_in, bookws_in,
                   booksort_in, moviews_in, moviesort_in, gamesws_in, gamesort_in, force=False):
    while True:
        if not force:
            if (local_in or (ip_in and username_in and key_in)) and file_in and level_in:
                if (bookws_in and booksort_in) or (not booksort_in and not bookws_in):
                    if(gamesws_in and gamesort_in) or (not gamesort_in and not gamesws_in):
                        if(moviews_in and moviesort_in) or (not moviesort_in and not moviews_in):
                            return [file_path, local_in, ip_in, username_in, key_in, file_in, level_in, token_in, 
                                    bookws_in, booksort_in, moviews_in, moviesort_in, gamesws_in, gamesort_in]
        print("Lets go through the settings")
        print("Where would you like the spreadsheet, logs, and information files stored")
        file_path_in = input("Enter the path to save files (C:\\Users\\Demo\\Desktop) [" + str(file_path_in) + "]: "
                             or file_path_in)
        print("Local mode lets everything run on this machine, if false we need an IP address of server, username,"
              " and key")
        local_in = bool(input("Local mode? True or False [" + str(local_in) + "]: ") or local_in)
        if not local_in:
            print("Since we are not in local mode, we will need the IP address of the server (ex 192.168.1.1)")
            ip_in = (input("Enter the servers IP address [" + str(ip_in) + "]: ") or ip_in)
            print("We will need the username that can be used to ssh into the server")
            username_in = (input("Enter the username [" + str(username_in) + "]: ") or username_in)
            print("We only allow a keyfile to ssh in, please provide the full path to the key")
            key_in = (input("Enter the path to the ssh key [" + str(key_in) + "]: ") or key_in)
        if local_in:
            print("If you would like, you can specify an IP address of a server in case "
                  "you want to switch out of local mode in the future")
            ip_in = (input("Enter the servers IP address [" + str(ip_in) + "]: ") or ip_in)
            print("We will need the username that can be used to ssh into the server")
            username_in = (input("Enter the username [" + str(username_in) + "]: ") or username_in)
            print("We only allow a keyfile to ssh in, please provide the full path to the key")
            key_in = (input("Enter the path to the ssh key [" + str(key_in) + "]: ") or key_in)
        print("We will also require the file name of the workbook, this is required even if uploading to"
              " dropbox")
        file_in = (input("Enter the excel workbook name [" + str(file_in) + "]: ") or file_in)
        print("We also require the log level to be used, if you aren't sure, type in 'INFO'")
        level_in = (input("Enter the log level [" + str(level_in) + "]: ") or level_in)
        print("If this database stores books, we will require the name of the worksheet, if you are not storing books, "
              "please leave blank")
        bookws_in = (input("If you will be storing books, enter the name of the Book Worksheet or leave blank [" +
                           str(bookws_in) + "]: ") or bookws_in)
        if bookws_in:
            print("Since we will be storing books, enter how you would like to sort books in reverse sort order "
                  "(start with the least significant column)")
            print("Possible columns are: Author, Title, Series, ISBN, Location, Condition, Borrowed By, Date Borrowed")
            if booksort_in:
                booksort_in = (input("Enter the sort order in a comma separated list in reverse order [" +
                                     ','.join(map(str, booksort_in)) + "]: ") or ','.join(map(str, booksort_in)))
            else:
                booksort_in = (input("Enter the sort order in a comma separated list in reverse order "
                                     "[Author,Series,Title]: ") or "Author,Series,Title")
        moviews_in = (input("If you will be storing movies, enter the name of the Movie Worksheet or leave blank [" +
                            str(moviews_in) + "]: ") or moviews_in)
        if moviews_in:
            print("Since we will be storing movies, enter how you would like to sort movies in reverse sort order "
                  "(start with the least significant column)")
            print("Possible columns are: Title, Series, UPC, Format, Borrowed By, Date Borrowed")
            if moviesort_in:
                moviesort_in = (input("Enter the sort order in a comma separated list in reverse order [" +
                                      ','.join(map(str, moviesort_in)) + "]: ") or ','.join(map(str, moviesort_in)))
            else:
                moviesort_in = (input("Enter the sort order in a comma separated list in reverse order "
                                      "[Title]: ") or "Title")
        gamesws_in = (input("If you will be storing games, enter the name of the Game Worksheet or leave blank [" +
                            str(gamesws_in) + "]: ") or gamesws_in)
        if gamesws_in:
            print("Since we will be storing games, enter how you would like to sort games in reverse sort order "
                  "(start with the least significant column)")
            print("Possible columns are: Title, Platform, UPC, Borrowed By, Date Borrowed")
            if gamesort_in:
                gamesort_in = (input("Enter the sort order in a comma separated list in reverse order [" +
                                     ','.join(map(str, gamesort_in)) + "]: ") or ','.join(map(str, gamesort_in)))
            else:
                gamesort_in = (input("Enter the sort order in a comma separated list in reverse order "
                                     "[Title,Format]: ") or "Title,Format")
        print("If you would like to upload the file to dropbox you will need to provide an access token, this is"
              " optional")
        token_in = (input("Enter the dropbox token [" + str(token_in) + "]: ") or token_in)

        settings_in = open("db.conf", "w")
        logging.info("Writing path = %s to db.conf", file_path_in)
        settings_in.write("path = " + str(file_path_in) + '\n')
        if file_path_in is not "":
            settings_in.close()
            settings_in = open(file_path+"db.conf", 'w')
        logging.info("Writing local = %s to db.conf", local_in)
        settings_in.write("local = " + str(local_in) + '\n')
        if bookws_in and booksort_in:
            logging.info("Writing bookWS = %s bookSort = %s to db.conf", bookws_in, ','.join(map(str, booksort_in)))
            settings_in.write("bookWS = " + bookws_in + '\n')
            settings_in.write("bookSort = " + booksort_in + '\n')
        if moviews_in and moviesort_in:
            logging.info("Writing movieWS = %s movieSort = %s to db.conf", moviews_in, ','.join(map(str, moviesort_in)))
            settings_in.write("movieWS = " + moviews_in + '\n')
            settings_in.write("movieSort = " + moviesort_in + '\n')
        if gamesws_in and gamesort_in:
            logging.info("Writing gameWS = %s gameSort = %s to db.conf", gamesws_in, ','.join(map(str, gamesort_in)))
            settings_in.write("gameWS = " + gamesws_in + '\n')
            settings_in.write("gameSort = " + gamesort_in + '\n')
        if ip_in:
            logging.info("Writing ip = %s to db.conf", ip_in)
            settings_in.write("ip = " + ip_in + '\n')
        if username_in:
            logging.info("Writing user = %s to db.conf", username_in)
            settings_in.write("user = " + username_in + '\n')
        if key_in:
            logging.info("Writing key = %s to db.conf", key_in)
            settings_in.write("key = " + key_in + '\n')
        if file_in:
            logging.info("Writing file = %s to db.conf", file_in)
            settings_in.write("file = " + file_in + '\n')
        if level_in:
            logging.info("Writing logLevel = %s to db.conf", level_in)
            settings_in.write("logLevel = " + level_in + '\n')
        if token_in:
            logging.info("Writing dropboxToken = %s to db.conf", token_in)
            settings_in.write("dropboxToken = " + token_in + '\n')
        settings_in.close()
        force = False


# Thread Manager monitors input queue and spawns sub threads to look up metadata. Auto saves every 10 inputs
def manage_thread(upc_queue, mt_wb, ssh_client):
    active_worksheet = mt_wb.ws
    my_logging = logging.getLogger("ThreadManager")
    my_logging.info("Thread Manager has started")
    blank = False
    # Used to track running sub threads
    running = []
    result = []
    previous = None
    # noinspection PyBroadException
    try:
        while True:
            my_logging.debug("Thread Manager is listening")
            # Blocking call for the next input
            entry = upc_queue.get()
            my_logging.debug("Received input %s", entry)
            # If the input is numeric assume it's a upc and do a search
            if entry.isnumeric():
                my_logging.info("Searching for %s", entry)
                # spawn a new thread to add book to database
                running.append(
                    future.ThreadPoolExecutor(thread_name_prefix='lookup ' + entry).submit(
                        active_worksheet.add_entry, ssh_client, entry))
            # if the input is the word exit then close all threads and exit
            elif entry == "exit":
                my_logging.debug("Waiting on metadata to finish")
                # wait on all threads in running to finish
                for j, x in enumerate(running):
                    while len(result) - 1 < j:
                        result.append(None)
                    result[j] = x.result()
                my_logging.debug("Checking for blanks before exiting")
                if active_worksheet.incomplete(result, previous):
                    blank = True
                my_logging.debug("Thread Manager is Exiting")
                return blank

            # After every new input check if the length of running is > 9 if it is
            if len(running) > 9:
                # save workbook before adding new threads
                my_logging.debug("Background Saving")
                for j, x in enumerate(running):
                    while len(result) - 1 < j:
                        result.append(None)
                    result[j] = x.result()
                my_logging.debug("Checking for blanks before continuing")
                if active_worksheet.ws.incomplete(result, previous):
                    blank = True
                previous = result[len(result) - 1]
                active_worksheet.save()
                running = []
                result = []
    except Exception as exc_inst:
        # Sanity check, manage Thread should never reach here
        my_logging.critical("manageThread left while loop, error is %s", exc_inst)
        print("Critical Error, manager has failed\nError is:\a", exc_inst)
        my_logging.critical("Traceback is %s", traceback.format_exc())
        exit()


# Always assume we will be connecting to a server
local = False
settings = settings2 = token = ip = username = key = file = level = bookWS = bookSort = movieWS = movieSort \
    = gameWS = gameSort = None
try:
    # Look for configuration File
    file_path = ""
    if not path.exists("db.conf"):
        file_path = input("could not find db.conf in local directory, please enter the path for where db.conf is "
                          "(C:\\Users\\Demo\\Desktop) (leave blank for current Directory: ")
        if file_path != "" and not file_path.endswith('\\'):
            file_path = file_path + '\\'
    settings = open(file_path+"db.conf", "r")
    temp = settings.readlines()
    # Checks for path to files to save
    file_path = next((s for s in temp if 'path' in s), file_path)
    if file_path != "":
        if '=' in file_path:
            file_path = file_path.split('=')[1].strip()
        my_logger.config_root_logger(file_path, 'INFO')
        logging.info("----------------------------------------------------------")
        logging.info("Files will be saved at %s", str(file_path))
        settings2 = open(file_path+"db.conf", "r")
        print("Read additional config from:", file_path+"db.conf")
        temp.extend(settings2.readlines())
    else:
        file_path = file_path.split('=')[1].strip()
        my_logger.config_root_logger(file_path, 'INFO')
        logging.info("----------------------------------------------------------")
    # Checks if using local mode
    local = next((s for s in temp if 'local' in s), None)
    if local is not None:
        local = bool(local.split('=')[1].strip())
        logging.info("Local mode is %s", str(local))
        print("Local mode is:", str(local))
    # Checks for IP address of the server to use (optional)
    ip = next((s for s in temp if 'ip' in s), None)
    if ip:
        ip = ip.split('=')[1].strip()
        logging.info("ip address is %s", ip)
        print("Connecting to", ip)
    # Username to SSH into server (optional)
    username = next((s for s in temp if 'user' in s), None)
    if username:
        username = username.split('=')[1].strip()
        logging.info("username is %s", username)
        print("as:", username)
    # Key to SSH into server (optional)
    key = next((s for s in temp if 'key' in s), None)
    if key:
        key = key.split('=')[1].strip()
        logging.info("key location is %s", key)
        print("using key at:", key)
    # Location to store database (required)
    file = next((s for s in temp if 'file' in s), None)
    if file:
        file = file.split('=')[1].strip()
        logging.info("Workbook is %s", file_path + file)
        print("using database at:" + file_path + file)
    # Logging Level (required)
    level = next((s for s in temp if 'logLevel' in s), None)
    if level:
        level = level.split('=')[1].strip()
        print("logging at:", level)
        logging.info("Setting log level to %s", level)
        my_logger.config_root_logger(file_path, level)
    # Dropbox Token (optional)
    token = next((s for s in temp if 'dropboxToken' in s), None)
    if token:
        token = token.split('=')[1].strip()
        print("Dropbox Token Found")
        logging.info("Found Access Token")
    bookWS = next((s for s in temp if 'bookWS' in s), None)
    if bookWS:
        bookWS = bookWS.split('=')[1].strip()
        print("Book Worksheet loaded:", bookWS)
        logging.info("Found Book Worksheet: %s", bookWS)
    bookSort = next((s for s in temp if 'bookSort' in s), None)
    if bookSort:
        bookSort = bookSort.split('=')[1].strip()
        print("Sorting books by:", bookSort)
        logging.info("Sorting books by: %s", bookSort)
        bookSort = [x.strip() for x in bookSort.split(',')]
        logging.debug("after splitting into an array: %s", bookSort)
    movieWS = next((s for s in temp if 'movieWS' in s), None)
    if movieWS:
        movieWS = movieWS.split('=')[1].strip()
        print("Movie Worksheet loaded:", movieWS)
        logging.info("Found Movie Worksheet: %s", movieWS)
    movieSort = next((s for s in temp if 'movieSort' in s), None)
    if movieSort:
        movieSort = movieSort.split('=')[1].strip()
        print("Sorting movies by:", movieSort)
        logging.info("Sorting books by: %s", movieSort)
        movieSort = [x.strip() for x in movieSort.split(',')]
        logging.debug("after splitting into an array: %s", movieSort)
    gameWS = next((s for s in temp if 'gameWS' in s), None)
    if gameWS:
        gameWS = gameWS.split('=')[1].strip()
        print("Game Worksheet loaded:", gameWS)
        logging.info("Found Game Worksheet: %s", gameWS)
    gameSort = next((s for s in temp if 'gameSort' in s), None)
    if gameSort:
        gameSort = gameSort.split('=')[1].strip()
        print("Sorting games by:", gameSort)
        logging.info("Sorting games by: %s", gameSort)
        gameSort = [x.strip() for x in gameSort.split(',')]
        logging.debug("after splitting into an array: %s", gameSort)
    settings.close()
    if settings2:
        settings2.close()
    file_path, local, ip, username, key, file, level, token, bookWS, bookSort, movieWS, movieSort = \
        check_settings(file_path, local, ip, username, key, file, level, token, bookWS, bookSort, movieWS, movieSort, 
                       gameWS, gameSort)
    my_logger.config_root_logger(file_path, level)

except OSError as inst:
    # If required information is absent, request all information and write it to settings file
    if settings:
        settings.close()

    logging.warning("Error reading information from file, error was %s", inst)
    print("Error reading information from file, error was "+inst.strerror)
    file_path, local, ip, username, key, file, level, token, bookWS, bookSort, movieWS, movieSort = \
        check_settings(file_path, local, ip, username, key, file, level, token, bookWS, bookSort, movieWS, movieSort, 
                       gameWS, gameSort)
    my_logger.config_root_logger(file_path, level)

if input("Settings appear to be valid, would you like to make any changes to them? yes/no [no]: ").lower() == "yes":
    file_path, local, ip, username, key, file, level, token, bookWS, bookSort, movieWS, movieSort = check_settings(
        file_path, local, ip, username, key, file, level, token, bookWS, bookSort, movieWS, movieSort, gameWS, 
        gameSort, True)
    my_logger.config_root_logger(file_path, level)

my_logger.clean_up_logs(file_path)
my_logger.start_thread_logging(file_path)
if not local:
    try:
        # If in remote mode, SSH into server
        logging.debug("Attempting to SSH into %s", ip)
        print("Connecting to", ip)
        client = paramiko.SSHClient()
        client.load_system_host_keys()
        client.set_missing_host_key_policy(paramiko.WarningPolicy)
        client.connect(ip, port=22, username=username, password="", pkey=None, key_filename=key)
    except Exception as inst:
        # If SSH fails for any reason fall back to local mode
        print("Unable to connect to host.")
        logging.critical("Failed to connect to %s using %s and %s error was:\n%s", ip, username, key, inst)
        print(inst)
        logging.critical("Falling back to local mode")
        print("Falling back to local mode\a")
        client = None
        local = True
else:
    # Set SSH client to none for local mode
    logging.debug("Skipping connecting to remote host")
    client = None
try:
    if token:
        # If Dropbox token is present, ask if we should download the spreadsheet or use a local one
        question = input("Dropbox token found, Download file?[yes]: ")
        if question.lower().strip() == "yes" or question.lower().strip() == "":
            logging.info("Attempting to download from dropbox")
            workbook.download_from_dropbox(file_path + file, token)
except Exception as inst:
    print("Failed to download %s from Dropbox, error was: %s", file, inst)
    logging.warning("Failed to download %s from Dropbox, error was: %s, trying to open local file", file, inst)
# Whether or not it was downloaded from dropbox, open the file locally
logging.info("Attempting to open file")
wb = workbook.MyWorkbook(file, file_path)
ws = None
print("Opened", file_path + file)
logging.info("Opened %s", file_path + file)
count = 0
for i, sheet in enumerate([bookWS, movieWS, gameWS]):
    count += 1
if count > 0:
    print("Database contains the following worksheets, enter the number of the worksheet to use")
    for i, sheet in enumerate([bookWS, movieWS, gameWS], start=1):
        print(str(i) + ") " + sheet)
    temp = input("Which Database? 1 - " + str(count)+": ")
    if temp == '1':
        logging.info("Opening Worksheet %s and sorting by %s", bookWS, bookSort)
        ws = wb.open_worksheet(bookWS, bookSort, book=True)
        sort = bookSort
    elif temp == '2':
        logging.info("Opening Worksheet %s and sorting by %s", movieWS, movieSort)
        ws = wb.open_worksheet(movieWS, movieSort, movie=True)
        sort = movieSort
    elif temp == '3':
        logging.info("Opening Worksheet %s and sorting by %s", gameWS, gameSort)
        ws = wb.open_worksheet(gameWS, gameSort, game=True)
if not ws:
    logging.critical("No valid worksheet, exiting")
    print("No valid worksheet, exiting")
    exit()
# Create the blocking queue and spawn the Thread Manager Thread
logging.info("Starting Thread Manager")
q = queue.Queue()
manager = future.ThreadPoolExecutor(thread_name_prefix='TM').submit(manage_thread, q, wb, client)
# Loop forever until exit or quit is typed in
while True:
    upc = input("Enter UPC: ")
    if upc.isnumeric():
        # If it's numeric put it in the queue for threadManager to use
        logging.debug("Searching for %s", upc)
        q.put(upc)
    elif upc == "exit" or upc == "quit":
        # Once told to exit, tell Thread Manager to wrap it up, then save and close the workbook before exiting
        print("waiting for metadata to finish downloading")
        logging.info("Waiting for metadata to finish downloading")
        q.put("exit")
        if manager.result():
            any_blanks = True
        else:
            any_blanks = False
        print("Saving Excel File")
        wb.save_and_close(token)
        if any_blanks:
            excel = None
            print("There was missing data.  Results have been saved to missing.txt")
            try:
                subprocess.Popen([r'C:\Program Files (x86)\Notepad++\notepad++.exe', file_path+'missing.txt'])
            except Exception as inst:
                print("Tried to open missing.txt using Notepad++ got error:", inst)
                logging.exception("Failed to open missing.txt error was %s", inst)
            try:
                excel = win32com.client.Dispatch("Excel.Application")

                logging.debug("Attempting to open %s", file_path + file)
                # Load file and Worksheet
                excel_wb = excel.Workbooks.Open(Filename=file_path + file)
                excel_ws = excel_wb.Worksheets(ws.name)
                excel.Visible = True
                print("Opened Excel, please make any changes and save and quit excel")
                while True:
                    # noinspection PyBroadException
                    try:
                        if excel.Visible:
                            logging.debug("Excel is active, sleeping for 5 seconds and checking again")
                            time.sleep(5)
                        else:
                            print("Excel is closed")
                            break
                    except:
                        print("Excel is closed")
                        break
                # noinspection PyBroadException
                try:
                    excel_wb.Save()
                except Exception as inst:
                    excel_wb = excel.Workbooks.Open(Filename=file_path + file)
                    excel_ws = excel_wb.Worksheets(ws.name)
                finally:
                    # Grab Table
                    tab = excel_ws.ListObjects("My"+ws.name)
                    # Remove current sorts, and sort by Author,Series,Title
                    tab.Sort.SortFields.Clear()
                    for sortValue in ws.sort_fields:
                        tab.Sort.SortFields.Add(Key=excel_ws.Range("My" + ws.name + "[" + sortValue + "]"), Order=1)
                    tab.Sort.Apply()
                    excel_wb.Save()
            except Exception as inst:
                print("Tried to open", file_path + file, "got error:", inst)
                logging.exception("Failed to open %s error was %s", file_path + file, inst)
            finally:
                if excel:
                    excel.Application.Quit()
                wb.upload_to_dropbox(token)
        if client:
            client.close()
        logging.info("Exiting Program")
        exit()
    else:
        # Otherwise give an error and wait for more input
        logging.warning("Error %s is not a valid upc", upc)
        print("UPC's are only numeric\a")
