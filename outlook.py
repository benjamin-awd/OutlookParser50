from tkinter import Tk
from tkinter.filedialog import askdirectory

import win32com.client
import sqlite3
import html

import re
import fnmatch
import os
from datetime import datetime
import logging

from setup_db import setup

# Setup database
db = setup()

def main():

    # Setup logger -- output to txt file and console
    logging.basicConfig(format='%(levelname)s:%(message)s', level=logging.DEBUG, 
                        handlers=[logging.FileHandler("log.txt"), logging.StreamHandler()])
    
    # Connect to Outlook by MAPI
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

    # Show dialog box and return folder path
    Tk().withdraw()
    folder_path = os.path.normpath(askdirectory(title='Select Folder'))

    # Init & populate list of emails
    email_list = [file for file in os.listdir(folder_path) if file.endswith(".msg")]

    # Get total emails by filtering for files in directory with the .msg extension
    email_total = len(fnmatch.filter(os.listdir(folder_path), '*.msg'))
    logging.info(f"Number of emails in directory = {email_total}")

    # Create counter to track total number of coverage
    total_coverage = 0

    # Iterate through every email
    for i, _ in enumerate(email_list):

        logging.info(email_list[i])

        # Create variable storing info from current email being parsed
        msg = outlook.OpenSharedItem(os.path.join(folder_path, email_list[i]))
        
        # Search email HTML for body text
        regex = re.search(r"<body([\s\S]*)</body>", msg.HTMLBody)
        body = regex.group()
        
        # Search body text for unique entries, as indicated by <li> tag TODO (add or statement for regex)
        # TRY li class=(MsoListParagraph|\"MsoListParagraph\")([\s\S]*?)</li> https://regex101.com/r/n530lx/1
        pattern = r"li class=(MsoListParagraph|\"MsoListParagraph\")([\s\S]*?)</li>"
        results = re.findall(pattern, body)
        num_coverage = len(results)

        # Check for alternate coding with quotes
        if num_coverage == 0:
            logging.warning(f"No coverage found in {email_list[i]}")

        # Keep track of total amount of coverage
        total_coverage += num_coverage

        logging.info(f"Processing email #{i + 1} out of {email_total}")
        logging.info(f"Coverage: {num_coverage} articles\n")

        # For each unique entry detected by regex, retrieve title, publication, pubtype and links based on HTML
        for header in results:
            logging.debug(header)
            
            # Get title, publication, platform, links from results
            title, publication, platform, links = get_title_pub(header)

            # Split string to parse variations in publication name e.g. HardwareZone vs HardwareZone Singapore
            pubsplit = publication.split()

            # Change date from "DD Month YY" to "dd/mm/yy"
            try:
                old_date_format = get_date(msg, publication, pubsplit)
                date = datetime.strptime(old_date_format, "%d %B %Y").strftime("%d/%m/%y")

            # If no dates are found, revert to send date
            except TypeError:
                date = msg.SentOn.strftime("%d/%m/%y")

            # In case of unknown date format
            except ValueError:
                pass

            # Get tier and category
            tier, category = get_tiercat(links, publication, pubsplit)

            logging.debug(f"Title - {title}")
            logging.debug(f"Pub - {publication}")
            logging.debug(f"Platform - {platform}")
            logging.debug(f"Link - {links}")
            logging.debug(f"Tier - {tier}")
            logging.debug(f"Category - {category}")
            logging.debug(f"Date - {date}\n")

            try:
                db.execute("INSERT INTO articles (date, title, publication, tier, category) VALUES (?, ?, ?, ?, ?)",
                           (date, title, publication, tier, category))

            # Where two articles have the same title, differentiate the second title by adding "(2)"
            except sqlite3.IntegrityError:
                title = title + "(2)"
                db.execute("INSERT INTO articles (date, title, publication, tier, category) VALUES (?, ?, ?, ?, ?)",
                           (date, title, publication, tier, category))

            # Copy id from main table into platforms and links table
            db.execute("INSERT INTO platforms (article_id) SELECT id FROM articles WHERE title = ?", (title,))
            db.execute("INSERT INTO links (article_id) SELECT id FROM articles WHERE title = ?", (title,))

            # Get article id and copy to platforms & links tables
            article_id = db.execute("SELECT id FROM articles WHERE title = ?", (title,))
            for item in article_id:
                _id = item[0]

            for i, _ in enumerate(platform):
                db.execute(f"UPDATE platforms SET platform{i} = ? WHERE article_id = ?", (platform[i], _id))

            for i, _ in enumerate(links):
                db.execute(f"UPDATE links SET link{i} = ? WHERE article_id = ?", (links[i], _id))

    db.commit()

    logging.info(f"Total coverage: {total_coverage}")


def get_title_pub(header):
    # Search HTML for title and publication, using regex to get text outside of <> and before )
    header = str(header)
    regex = re.search(r"[^<>]+(?=\(|sans-serif'>([\s\S]*?)</span>)", header)

    # HTML unescape to get rid of extant HTML like unicode dash '&8211;'
    title_pub = html.unescape(regex.group())

    # Create variable storing split strings
    _list = title_pub.split('–')

    # If no split occured, try splitting by short dash '-' instead
    if len(_list) < 2:
        _list = title_pub.split('-')
        # If still cannot split, return error
        if len(_list) < 2:
            logging.debug(f"\ntitle_pub = {title_pub}")
            logging.warning("Error: could not retrieve title/publication")

    # Create empty list for pubtype variable
    platform = []

    # Check 1: Search for regular expression
    regex = re.findall(r"[^\">]+(?=</a)", header)
    if len(regex) > 0:
        for match in regex:
            match = html.unescape(match)
            platform.append(match)

    # Check 2: compare strings in header against various pub platforms
    pub_typelist = ["Online", "Facebook", "Instagram", "Twitter", "LinkedIn", "Linkedin", "Youtube"]
    if len(platform) == 0:
        for pub in pub_typelist:
            if pub in header:
                platform.append(pub)
                logging.debug("Match found!")
                logging.debug(f"platform = {pub}")

    # Check for print publications
    print_check = re.findall(r"\bPrint\b", header)
    if print_check != None:
        for i in print_check:
            print_type = i
            platform.append(print_type)

    # Find all links using regex
    links = re.findall(r"<a href=\"([\s\S]*?)\">", header)

    # Remove blank spaces in title and publication & record as variables
    try:
        # Account for extra dash in title, assuming that the last split string is the publication name
        if len(_list) > 2:
            title = _list[1].strip()
            publication = _list[2].strip()

        else:
            title = _list[0].strip()
            publication = _list[1].strip()

    except IndexError:
        title = _list[0].strip()

    try:
        return title, publication, platform, links

    except UnboundLocalError:
        publication = "N/A"
        return title, publication, platform, links


def get_date(msg, publication, pubsplit):
    count_attachments = msg.Attachments.Count

    # Convert first 4 chars of send date into a string
    year = str(msg.SentOn)[0:4]

    # If there are attachments, iterate through each attachment and compare publication name against publication name in attachment filename
    if count_attachments > 0:
        for item in range(count_attachments):
            filename = msg.Attachments.Item(item + 1).Filename
            if (re.search(f"^.*{publication}.*$", filename, flags=re.I) != None):
                year = str(msg.SentOn)[0:4]
                date = f"{filename.split('-')[0]}{year}"
                return date

        # If loop completes and no exact match found, look for alternative
        logging.debug("No exact match found -- looking for alternative match")
        for item in range(count_attachments):

            # Retrieve filename of attachment
            filename = msg.Attachments.Item(item + 1).Filename

            for i in range(len(pubsplit) - 1):
                # Check if filename matches first word in publication name
                logging.debug(f"Checking for '{pubsplit[i]}' in {filename} . . .", end=" ")
                if (re.search(f"^.*{pubsplit[i]}.*$", filename, flags=re.I) != None):
                    if len(pubsplit) == 1:
                        date = f"{filename.split('-')[0]}{year}"
                        print("Match found!")
                        return date

                    logging.debug(f"Checking for '{pubsplit[i+1]}' in {filename} . . .", end=" ")
                    # Check if filename matches second word in publication name
                    if (re.search(f"^.*{pubsplit[i+1]}.*$", filename, flags=re.I) != None):

                        # Return date if two matches found
                        logging.debug("Match found!")
                        date = f"{filename.split('-')[0]}{year}"
                        return date

                # If both checks fail, check for matches of first word without special chars
                else:
                    regex = re.compile('[^a-zA-Z]')
                    new_string = regex.sub("", pubsplit[i])
                    if (re.search(f"^.*{new_string}.*$", filename, flags=re.I) != None):
                        date = f"{filename.split('-')[0]}{year}"
                        return date
        

        logging.warning("Error: could not retrieve date from attachment, reverting to send date")
        date = str(msg.SentOn)[0:9]
        return date


def get_tiercat(links, publication, pubsplit):
    database = db.execute("SELECT * FROM medialist")

    # Get maximum rows
    count = db.execute("SELECT COUNT(*) FROM medialist")
    for result in count:
        max_rows = result[0]

    # Remove spaces from publication name
    string = publication.replace(" ", "")

    # Create variable to keep track of rows
    row_counter = 0

    for row in database:
        row_counter += 1
        url = row[0]

        # Create a variable that stores the initial part of the link e.g. www.hardwarezone.com.sg
        regex = re.search(r"(?<=//)(.*?)(?=\/)", links[0])
        link = regex.group()

        # Check 1: Compares the database URL and the link URL with each other - if match found, return tier and media type
        if link.find(f"{url}") != -1 or url.find(f"{link}") != -1:
            tier = row[1]
            category = row[2]
            return tier, category

        # Check 2: Compare publication name and truncated version of publication name against database
        if row_counter == max_rows:
            database = db.execute("SELECT * FROM medialist")
            row_counter = 0

            for row in database:
                row_counter += 1
                url1 = row[0]
                url2 = row[0].replace(" ", "")
                logging.debug(f"Checking for {string} against {url1}")
                if re.search(f"{string}", url1, flags=re.I) != None or re.search(f"{string}", url2, flags=re.I) != None:
                    tier = row[1]
                    category = row[2]
                    return tier, category

        # Check 3: take the first word of the publication name from email and compare against database
        if row_counter == max_rows:
            database = db.execute("SELECT * FROM medialist")
            row_counter = 0

            for row in database:
                row_counter += 1
                url1 = row[0]
                logging.debug(f"Checking for '{pubsplit[0]}' in {url1} . . .")
                if (re.search(f"{pubsplit[0]}", url1, flags=re.I) != None):
                    tier = row[1]
                    category = row[2]
                    return tier, category

        # Check 4: take the first word of the publication name from email, remove non-alphanumeric chars then compare against database
        if row_counter == max_rows:
            database = db.execute("SELECT * FROM medialist")
            row_counter = 0

            regex = re.compile('[^a-zA-Z]')
            new_string = regex.sub("", pubsplit[0])
            if new_string != "":
                for row in database:
                    row_counter += 1
                    url1 = row[0]
                    logging.debug(f"Checking for '{new_string}' in {url1} . . .")
                    if (re.search(f"{new_string}", url1, flags=re.I) != None):
                        tier = row[1]
                        category = row[2]
                        return tier, category

    logging.warning("Tier/Type not found")
    tier = category = ("N/A")
    return tier, category

main()