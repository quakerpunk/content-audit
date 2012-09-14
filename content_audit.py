#!/usr/bin/env python
# content_audit.py
# Python script that grabs data from a given website.

from bs4 import BeautifulSoup
from optparse import OptionParser
import urllib2
import random
import xlwt
import urlparse
import re
import time
import logging

class ContentAuditor:
    """
    ContentAuditor

    This script takes a list of URLs and retrieves data such as meta tags
    containing keywords, description and the page title. It then populates a
    spreadsheet with the data for easy review.
    """

    def __init__(self, filename):
        """
        Initialization method for the ContentAuditor class.

        requirements:
        BeautifulSoup for HTML parsing
        xlwt for writing to an Excel spreadsheet without use of COM Interop
        """
        self.filehandle = open(filename, 'r')
        self.soupy_data = ""
        self.workbook = ""
        self.current_sheet = ""
        self.content = ""
        self.text_file = ""
        self.site_info = []
        self.url_parts = ""
        self.reg_expres = re.compile(r"www.(.+?)(.com|.net|.org)")

    def read_url(self):
        """
        read_url

        Method which reads in a given url (to the constructor) and puts data
        into a BeautifulSoup context.

        We begin setting a string for the user-agent. Checking for comment
        lines in the URL list, we take a web address, one at a time, download
        the HTML, parse it with BeautifulSoup then pass it off to extract tags.
        Along the way, we check for any connectivity or remote server issues
        and handle them appropriately.
        """
        right_now = str(int(time.time()))
        errlogname = 'content_audit_log_' + right_now
        logging.basicConfig(filename=errlogname, level=logging.INFO, format='%(levelname)s: %(message)s')
        logging.info('Beginning extraction...\n')
        ua_string = 'Content-Audit/2.0'
        for line in self.filehandle:
            if line.startswith("#"):
                continue
            print "Parsing %s" % line
            self.url_parts = urlparse.urlparse(line)
            req = urllib2.Request(line)
            req.add_header('User-Agent', ua_string)
            try:
                data = urllib2.urlopen(req)
            except urllib2.HTTPError, ex:
                logging.warning("Could not parse %s", line.rstrip())
                logging.warning("The server returned the following: ")
                logging.warning("Error code: %s", ex.code)
                logging.warning("Moving on to the next one...\n")
                continue
            except urllib2.URLError, urlex:
                logging.warning("Could not parse %s", line.rstrip())
                logging.warning("We did not reach a server.")
                logging.warning("Reason: %s", urlex.reason)
                logging.warning("Moving on to the next one...\n")
                continue
            self.soupy_data = BeautifulSoup(data)
            self.extract_tags()
            time.sleep(random.uniform(1, 3))
        logging.info("End of extraction")

    #Extraction methods

    def extract_tags(self):
        """
        extract_tags

        Searches through self.soupy_data and extracts meta tags such as page
        description and title for inclusion into content audit spreadsheet
        """
        page_info = {}

        for tag in self.soupy_data.find_all('meta', attrs={"name": True}):
            page_info[tag['name']] = tag['content']
        page_info['title'] = self.soupy_data.head.title.contents[0]
        page_info['filename'] = self.url_parts[2]
        page_info['name'] = self.soupy_data.h3.get_text()
        self.add_necessary_tags(page_info, ['keywords', 'description', 'title'])
        self.site_info.append(page_info)
        self.soupy_data = ""

    #Spreadsheet methods

    def write_to_spreadsheet(self):
        """
        write_to_spreadsheet

        Write data from self.meta_info to spreadsheet. Worksheet takes name of
        url
        """
        self.workbook = xlwt.Workbook()

        page_name = self.reg_expres.match(self.url_parts[1])
        self.current_sheet = self.workbook.add_sheet(page_name.group(1))

        self.current_sheet.write(0, 0, "Page Name")
        self.current_sheet.write(0, 1, "File Name")
        self.current_sheet.write(0, 2, "Page Title")
        self.current_sheet.write(0, 3, "Page Description")
        self.current_sheet.write(0, 4, "Keywords")
        self.current_sheet.write(0, 5, "Notes")

        count = 1

        for dex in self.site_info:
            self.current_sheet.write(count, 0, dex['name'])
            self.current_sheet.write(count, 1, dex['filename'])
            self.current_sheet.write(count, 2, dex['title'])
            self.current_sheet.write(count, 3, dex['description'])
            self.current_sheet.write(count, 4, dex['keywords'])
            count += 1

        #self.workbook.save('content_audit.xls')
        self.workbook.save(options.output)

    #Helper methods

    def add_necessary_tags(self, info_dict, needed_tags):
        """
        add_necessary_tags

        This method insures that missing tags have a null value
        before they are written to the output spreadhseet.
        """
        for key in needed_tags:
            if key not in info_dict:
                info_dict[key] = " "
        return info_dict

if __name__ == "__main__":
    parser = OptionParser()
    parser.add_option("-f", "--file", dest="filename",
                      help="Filename containing URLs", metavar="FILE")
    parser.add_option("-o", "--output", dest="output",
                      help="Output file, usually a spreadsheet")

    (options, args) = parser.parse_args()

    if not options.output:
        parser.error("You did not specify an output file.")

    if options.filename:
        content_bot = ContentAuditor(options.filename)
        content_bot.read_url()
        content_bot.write_to_spreadsheet()
    else:
        parser.error("You did not specify an input file (a list of URLs)")
