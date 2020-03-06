import requests
import logging
import re


def long_substr(data):
    substr = ''
    if len(data) > 1 and len(data[0]) > 0:
        for i in range(len(data[0])):
            for j in range(len(data[0])-i+1):
                if j > len(substr) and all(data[0][i:i+j] in x for x in data):
                    substr = data[0][i:i+j]
    return substr


def search_imdb(imdb_id):
    url = "http://www.imdb.com/"+imdb_id
    logging.debug("url: %s", url)
    response1 = requests.request("GET", url)
    logging.debug("url: %s", url+"movieconnections/?tab=mc&ref_=tt_trv_cnn")
    response2 = requests.request("GET", url+"movieconnections/?tab=mc&ref_=tt_trv_cnn")
    return [response1, response2]


def search_moodb(upc):
    url = "http://www.moodb.net/search.asp"
    response2 = imdb_id = None
    logging.debug("Searching: %s", url)
    payload = "cboMainSearchType=M&txtUPC=" + upc + "&btnSearchMovies=Search%20for%20movies"
    headers = {
        'Host': "www.moodb.net",
        'Connection': "keep-alive",
        'Cache-Control': "max-age=0",
        'Origin': "http://www.moodb.net",
        'Upgrade-Insecure-Requests': "1",
        'DNT': "1",
        'Content-Type': "application/x-www-form-urlencoded",
        'Accept': "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;"
                  "q=0.8,application/signed-exchange;v=b3;q=0.9",
        'Accept-Encoding': "gzip,deflate",
        'Accept-Language': "en-US,en;q=0.9",
        'Content-Length': "76",
        'User-Agent': "Mozilla%20/%205.0(Windows%20NT%2010.0;%20Win64;%20x64)%20AppleWebKit%20/%20537.36"
                      "(KHTML,%20like%20Gecko)%20Chrome%20/%2080.0.3987.122%20Safari%20/%20537.36"
    }
    my_cookie = {"Cookie": "ASPSESSIONIDSQATQQSS=PJOFIJPCNFOOOPJPPFPNGCLO"}
    logging.debug("url: %s, payload: %s, headers: %s, cookie: %s", url, payload, headers, my_cookie)
    response1 = requests.request("POST", url, data=payload, headers=headers, cookies=my_cookie)
    match1 = re.search(r'<a href=(.*?)</a><br>', str(response1.text))
    if match1 is not None:
        match1 = match1.group(1)
        movie_id = re.search(r'id=(.*?)"', str(match1)).group(1)
        logging.debug("hit: %s", match1)
        logging.debug("Doing second search on moodb for additional information")
        url = "http://www.moodb.net/movie.asp?id=" + movie_id
        headers = {
            'Host': "www.moodb.net",
            'Connection': "keep-alive",
            'Upgrade-Insecure-Requests': "1",
            'DNT': "1",
            'User-Agent': "Mozilla%20/%205.0(Windows%20NT%2010.0;%20Win64;%20x64)%20AppleWebKit%20/%20537.36"
                          "(KHTML,%20like%20Gecko)%20Chrome%20/%2080.0.3987.122%20Safari%20/%20537.36",
            'Accept': "text/html",
            'Referrer': "http://www.moodb.net/searchresult.asp?mainsearch",
            'Accept-Encoding': "gzip,deflate",
            'Accept-Language': "en-US,en;q=0.9"
        }
        logging.debug("Second request url: %s, headers: %s", url, headers)
        response2 = requests.request("GET", url, headers=headers, cookies=my_cookie)
        match2 = re.search(r'<a href="http://www.imdb.com/(.*?)">', str(response2.text))
        if match2 is not None:
            imdb_id = match2.group(1)
    return [response1, response2, imdb_id]


def search_upc_scavenger(upc):
    url = "http://www.upcscavenger.com/barcode/" + upc
    logging.debug("Searching: %s", url)
    headers = {
        'User-Agent': "Mozilla%20/%205.0(Windows%20NT%2010.0;%20Win64;%20x64)%20AppleWebKit%20/%20537.36"
                      "(KHTML,%20like%20Gecko)%20Chrome%20/%2080.0.3987.122%20Safari%20/%20537.36",
        'Accept': "*/*",
        'Cache-Control': "no-cache",
        'Host': "www.upcscavenger.com",
        'Accept-Encoding': "gzip, deflate",
        'Connection': "keep-alive"
    }
    logging.debug("url: %s, headers: %s", url, headers)
    response = requests.request("GET", url, headers=headers)
    return response


def get_metadata(upc):
    logging.info("Starting Search for %s", upc)
    response = [None, None, None, None, None]
    response[0], response[1], imdb_id = search_moodb(upc)
    if imdb_id is not None:
        response[2], response[3] = search_imdb(imdb_id)
    response[4] = search_upc_scavenger(upc)
    metadata = [None, None, None, upc]
    metadata[0], metadata[1], metadata[2] = parse_metadata(response)
    logging.debug("Displaying resposnes from web search for Movie:\n\nMoodb first page:\n%s\n\nMoodb second page:\n%s"
                  "\n\nIMDB first page:\n%s\n\nIMDB second page:\n%s\n\nUPC Scavenger:\n%s", response[0], response[1],
                  response[2], response[3], response[4])
    logging.info("Finished searching for %s", upc)
    return metadata


def parse_metadata(response):
    metadata_out = [None, None, None]
    logging.debug("Parsing responses for metadata")

    if response[2] is not None:
        # Check IMDB for the title
        imdb_search = re.search(r'<title>(.*?)</title>', str(response[2].text))
        if imdb_search is not None:
            imdb_search = imdb_search.group(1)
            if imdb_search.rfind('-') > 0:
                imdb_search = imdb_search[:imdb_search.rfind('-')]
            if imdb_search.rfind('(') > 0:
                imdb_search = imdb_search[:imdb_search.rfind('(')]
            metadata_out[0] = imdb_search.strip()

    if response[3] is not None:
        # Check IMDB for any Series Information
        imdb_search = re.search(r'<a id="follows"(.*?)<a id=', str(response[3].text), re.DOTALL)
        movies = []
        if imdb_search is not None:
            movie_list = re.findall(r'<a href="(.*?)</a>', imdb_search.group(1))
            for movie in movie_list:
                movies.append(re.search(r'">(.*?)</', movie))
        imdb_search = re.search(r'<a id="followed_by"(.*?)<a id=', str(response[3].text), re.DOTALL)
        if imdb_search is not None:
            movie_list = re.findall(r'<a href="(.*?)</a>', imdb_search.group(1))
            for movie in movie_list:
                movies.append(re.search(r'">(.*?)</', movie))
        if len(movies) > 1:
            metadata_out[1] = long_substr(movies)

    if response[0] is not None:
        # Checking moodb's initial search for the title
        if metadata_out[0] is None:
            moodb_search = re.search(r'<a href=(.*?)</a><br>', str(response[0].text))
            if moodb_search is not None:
                moodb_search = moodb_search.group(1)
                metadata_out[0] = moodb_search[moodb_search.rfind('>') + 1:]
                logging.debug("Moodb got title: %s", metadata_out[0])

    if response[4] is not None:
        # If no title, check UPC Scavenger
        if metadata_out[0] is None:
            logging.debug("Moodb didn't get title, checking UPC Scavenger")
            upc_search = re.search(r'<title>(.*?)</title>', str(response[4].text))
            if upc_search is not None:
                upc_search = upc_search.group(1)
                upc_search = upc_search[upc_search.rfind(':') + 1:].strip()
                if upc_search.rfind('-') > 0:
                    upc_search = upc_search[:upc_search.rfind('-')]
                if upc_search.rfind('(') > 0:
                    upc_search = upc_search[:upc_search.rfind('(')]
                metadata_out[0] = upc_search.strip()
                logging.debug("UPC Search got result: %s", metadata_out[0])

        # Check UPC Scavenger for format
        if metadata_out[2] is None:
            upc_search = re.search(r'<title>(.*?)</title>', str(response[4].text))
            if upc_search is not None:
                upc_search = upc_search.group(1)
                if upc_search.lower().find('dvd') > -1:
                    metadata_out[2] = "DVD"
                    logging.debug("UPC Search found DVD inside %s", upc_search)
                if upc_search.lower().find('bluray') > -1 or upc_search.lower().find('blu-ray') > -1:
                    metadata_out[2] = "Blu-ray"
                    logging.debug("UPC Search found Blu-ray inside %s", upc_search)
            # Title had no results, checking descriptions for format
            if metadata_out[2] is None:
                upc_search = re.findall(r'description"(.*?)">', response[4].text, re.DOTALL)
                if upc_search is not None:
                    logging.debug("UPC Search did not find format in the title, checking description for format")
                    for result in upc_search:
                        if result.lower().find('dvd') > -1:
                            metadata_out[2] = "DVD"
                            logging.debug("UPC Search found DVD inside %s", result)
                            break
                        if result.lower().find('bluray') > -1 or result.lower().find('blu-ray') > -1:
                            metadata_out[2] = "Blu-ray"
                            logging.debug("UPC Search found Blu-ray inside %s", result)
                            break

    return metadata_out
