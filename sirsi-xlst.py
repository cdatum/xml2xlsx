#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Exporting Sirsi XML report to Excel
"""
from bs4 import BeautifulSoup
import xlsxwriter
from datetime import datetime  # for date/timestamp
import os # for os.listdir()


'''
# Some data we want to write to the worksheet.
expenses = (
    ['Rent', 1000],
    ['Gas',   100],
    ['Food',  300],
    ['Gym',    50],
)

# Start from the first cell. Rows and columns are zero indexed.
row = 0
col = 0

# Iterate over the data and write it out row by row.
for item, cost in (expenses):
    worksheet.write(row, col,     item)
    worksheet.write(row, col + 1, cost)
    row += 1

# Write a total using a formula.
worksheet.write(row, 0, 'Total')
worksheet.write(row, 1, '=SUM(B1:B4)')

workbook.close()
'''



# Holds the directory where xml file(s) will be sourced and docx file(s) will be saved
directory_path = ''
# Flag for when no XML files are found
no_xml = True

def get_catalog_details(tag):
    print("name:", tag.name)
    print("attrs:", tag.attrs)

# Get a list of .xml files in the directory
def get_filelist():
    xml_filenames = []
    # https://docs.python.org/3/library/os.html#os.listdir
    # A GUI might be nice for finding the right directory
    dir = input("Enter a directory that contains xml files to be converted OR hit enter to continue with current dir. \n")
    
    global directory_path
    global no_xml
    # Use the current directory if the user inputs nothing
    if dir == '':
        directory_path = os.getcwd()
        files_list = os.listdir()
        print("Searching", os.getcwd()) # https://docs.python.org/3/library/os.html#os.getcwd
        # Otherwise use the user-specified directory
    else:
        directory_path = dir
        files_list = os.listdir(dir)
        print("Searching", dir)
    
    for file in files_list:
        if file.endswith('.xml'):
            xml_filenames.append(file)
            
    
    # Let user know how many files were found
    if len(xml_filenames) == 0:
        print("\nSorry, no files found in that directory")
        
    elif len(files_list) == 1:
        print("\nFound", len(xml_filenames), "file to convert.")
        no_xml = False
    else:
        print("\nFound", len(xml_filenames), "files to convert. ")
        no_xml = False
    

    # Iterate through each file, process it, and generate a Word doc
    for xml_file in xml_filenames:          
        convert_xml_to_word(xml_file)    
    

# This function extracts the bulk of the data from the xml file.
# Accepts an xml file and produces a Word file. Only prints select
# fields; not an entire bib or item record
def convert_xml_to_word(xml_file):
    

    # Grab the source filename minus '.xml' This will be used in the title of the .docx
    file_title = xml_file.replace('.xml', '') + '_.xlsx'
    
    # Create a workbook and add a worksheet.
    workbook = xlsxwriter.Workbook(file_title)
    worksheet = workbook.add_worksheet()

    

    #doc.add_heading(str(datestamp), level=3)
    print("\nConverting", xml_file)
    
    with open(xml_file, encoding = 'utf-8') as booklist:
        global directory_path
        # a list to hold circ stats
        stats_list = []
        
        # a list to hold bib details
        item_list = []
        
        soup  = BeautifulSoup(booklist, 'lxml-xml')
        marc = soup.find_all('catalog')
        tag = soup.marcEntry
        # should probably handle an xml file that isn't from Sirisi or one that isn't formatted properly
        count = 1
        for item in marc:
            
            # find_all() returns a _list_ of tags and strings–a ResultSet object. 
            # You need to iterate over the list and look at the .foo of each one.
            # https://www.crummy.com/software/BeautifulSoup/bs4/doc/
            if item.yearOfPublication != None:
                year = item.yearOfPublication.get_text()
            else: 
                year = "Check year" # A few titles like yearbooks don't have a year!
                
            barcode = item.itemID.get_text()   
            total_charges = item.totalCharges.get_text()
            date_last_use = item.dateLastUsed.get_text()
            call_no = item.callNumber.get_text()
            isbn = ''
            # This find_all returns a list of all the marcEntry elements
            tag = item.find_all('marcEntry')

            
            # list to hold details of one item
            item_details = []
            for element in tag: 
                # Title
                # if 245 in element[tag] 
                if element['tag'] == '245':
                    title = element.get_text()
                    stats_title = element.get_text()
                    title = str(count) + ". " + element.get_text() + " (" + year + ")"
                    
                # Description
                if element['tag'] == '260':
                    desc = element.get_text()
                    
                # Physical desc
                if element['tag'] == '300':                    
                    desc += " " + element.get_text()
                    
                # Campus Location    
                if element['tag'] == '596':
                    campus = element.get_text()
                
                # ISBNS
                if element['tag'] == '020':
                    num  = element.get_text() + ' | '
                    isbn += num
            
            # Add bibliographic details to Excel file 
            item_details.extend([title,barcode,campus,call_no,isbn,desc,total_charges,date_last_use])
            
            item_list.append(item_details)
            
            # Collect some basic data for item stats 
            item_stats = []
            elements = (barcode, stats_title, total_charges)
            item_stats.extend(elements)
            # Save item stats in a list for the top ten list
            stats_list.append(item_stats)
            
            count += 1
        
        '''

        # Display some stats for the user to see while program is running
        print(str(count), "Total items\t", get_circ_stats(stats_list), "Total checkouts")
        
        # Add date and brief summary of items        
        date_paragraph.insert_paragraph_before("Report created on " + soup.dateCreated.get_text())
        
        number_of_items = [str(count), "total items in this report.\t" , str(get_circ_stats(stats_list)), " total checkouts."]
        subtitle_data = ' '        
        date_paragraph.insert_paragraph_before(subtitle_data.join(number_of_items))
        
        # Add table at the end of the report that displays top 10 items that circulated
        get_top_ten(stats_list, doc)
        
'''          
        
        print("Done!")
     # Start from the first cell. Rows and columns are zero indexed.  
        row = 0
        col = 0            

    # Iterate over the data and write it out row by row.
        for item in (item_list[row]):
            worksheet.write(row, col, item[0])
            worksheet.write(row, col + 1, item[1])
            worksheet.write(row, col + 1, item[2])
            worksheet.write(row, col + 1, item[3])
            worksheet.write(row, col + 1, item[4])
            worksheet.write(row, col + 1, item[5])
            
            row += 1            
                
        workbook.close() 

        
        
        
# Returns a total # of checkouts for all titles within an XML file
# Accepts a list of items as a parameter; returns an int of total charges (checkouts + renewals)
def get_circ_stats(item_list):
    sort_list = []
    sort_list = sorted(item_list, key=lambda record: int(record[2]), reverse=True)
    # from https://docs.python.org/3/howto/sorting.html#sortinghowto

    charges = 0
    for item in sort_list:
        charges += int(item[2])

    return charges


    
    
# Get a timestap and use it to create a unique filename    
def get_file_extension():
    date_extension = datetime.today()
    date_extension = str(date_extension).replace('-','_')
    date_extension = date_extension.replace(':', '')
    date_extension = date_extension.replace(' ', '-')
    date_extension = date_extension.replace('.', '')
    date_extension = date_extension + '.docx'
    return date_extension


def say_goodbye():
    if no_xml == False:
        print("\nAll done! \n\nFind the Word file(s)in the following directory:\n", directory_path )
    else:
        print("\nLet's try again" )
        get_filelist()
        #TODO add a while loop for better UX 
                

#Let user enter the name of a directory
file_list = get_filelist()

say_goodbye()


