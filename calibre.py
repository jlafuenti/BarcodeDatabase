import logging
import subprocess
import re


# Parses the metadata into an array, formats the authors 'Lastname, Firstname'
def parse_metadata(stdout, logger):
    logger.debug("Parsing Metadata")
    # Decode it in utf-8 splitting on the new lines
    parsed = stdout.decode("utf-8").split('\n')
    logger.debug(parsed)
    output = ["", "", ""]
    # Find the Author Field and remove extra characters
    output[0] = [x for x in parsed if re.search('Author', x)]
    if len(output[0]) > 0:
        output[0] = output[0][0].split(':', 1)[1].strip()
        # If there are more than one author, handle each individually
        temp = output[0].split('&')
        for i, s in enumerate(temp):
            temp[i] = temp[i].strip()
            # If the author has ' Del ' then count that as part of the last name
            # Either way, put the last name before the first name separated by a comma
            if s.find(' Del ') != -1:
                temp[i] = temp[i][temp[i].rfind(' Del '):] + ', ' + temp[i][0:temp[i].rfind(' Del ')]
            else:
                if len(temp[i].split(' ')) > 1:
                    temp[i] = temp[i][temp[i].rfind(' '):] + ', ' + temp[i][0:temp[i].rfind(' ')]
            temp[i] = temp[i].strip()
        # Sort the lastnames for each author alphabetically
        temp.sort()
        # Rejoin them together into a string
        s = " & "
        s = s.join(temp)
        output[0] = s
    else:
        output[0] = None
    logger.debug("Author:%s", output[0])
    # Find the Title field and remove extra characters
    output[1] = [x for x in parsed if re.search('Title', x)]
    if len(output[1]) > 0:
        output[1] = output[1][0].split(':', 1)[1].strip()
    else:
        output[1] = None
    logger.debug("Title:%s", output[1])
    # Find the Series field and remove extra characters
    output[2] = [x for x in parsed if re.search('Series', x)]
    if len(output[2]) > 0:
        # Sometimes they download with 'Serie' in their, if it is present remove it
        output[2] = output[2][0].split(':', 1)[1].strip().replace(" Serie ", " ").strip()
    else:
        output[2] = None
    logger.debug("Series:%s", output[2])
    # Return the array of parsed metadata
    return output


# Download the metadata for the isbn provided
def get_metadata(client, isbn):
    logger = logging.getLogger("Calibre")
    logger.info("Getting Metadata for %s", isbn)
    # If there is a remote server running calibre in docker, have the server download the metadata
    if client:
        stdin, stdout, stderr = client.exec_command('docker exec -t calibre fetch-ebook-metadata --isbn ' + isbn)
        stdout.channel.recv_exit_status()
        logger.debug("Parsing Metadata")
        # Pass the output to be parsed
        metadata = parse_metadata(stdout.read(), logger)
    else:
        # If in local mode, get the metadata directly
        output = subprocess.run(["fetch-ebook-metadata", "--isbn", isbn], stdout=subprocess.PIPE,
                                stderr=subprocess.PIPE)
        # Pass the output to be parsed)
        metadata = parse_metadata(output.stdout, logger)
    # Add the ISBN to the parsed metadata and return it
    metadata.append(isbn)
    logger.info("Finished searching for %s", isbn)
    return metadata
