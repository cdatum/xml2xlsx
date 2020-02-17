#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Exporting Sirsi XML report to Excel
"""
from bs4 import BeautifulSoup
import xlsxwriter
import os # for os.listdir()


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
        convert_xml_to_xlsx(xml_file)    
    

# This function extracts the bulk of the data from the xml file.
# Accepts an xml file and produces an Excel file. Only prints select
# fields; not an entire bib or item record
def convert_xml_to_xlsx(xml_file):    

    # Grab the source filename minus '.xml' This will be used in the title of the resulting file
    file_title = xml_file.replace('.xml', '') + '_.xlsx'
    
    # Create a workbook and add a worksheet.
    workbook = xlsxwriter.Workbook(file_title)
    worksheet = workbook.add_worksheet()

        
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
                    title = element.get_text() + " (" + year + ")"
                    
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
        
        
        
        print("Done!")
        

        
        # Start from the first cell. Rows and columns are zero indexed.  
        row = 0
        col = 0            
        
        # Setup col & cell formatting
        header_format = workbook.add_format({'bold': True})
        text_wrap_format = workbook.add_format()
        text_wrap_format.set_text_wrap()
        worksheet.set_column('A:A' , 60)  # Width title col
        worksheet.set_column('B:E' , 20)    # Width cols B-E
        worksheet.set_column('F:F' , 60)    # Width desc col
        
        # Output the results starting with row headers
        worksheet.write(row, col, 'Title', header_format)
        worksheet.write(row, 1, 'Barcode', header_format)
        worksheet.write(row, col + 2, 'Campus', header_format)
        worksheet.write(row, col + 3, 'Call Number', header_format)
        worksheet.write(row, col + 4, 'ISBN', header_format)
        worksheet.write(row, col + 5, 'Description', header_format)
        worksheet.write(row, col + 6, 'Total Charges', header_format)
        worksheet.write(row, col + 7, 'Date of Last Use', header_format)
        
        row += 1 

        # Iterate over the data and write it out row by row.
        for item in item_list:
            worksheet.write(row, col, item[0], text_wrap_format)
            worksheet.write(row, col + 1, item[1])
            worksheet.write(row, col + 2, item[2])
            worksheet.write(row, col + 3, item[3])
            worksheet.write(row, col + 4, item[4], text_wrap_format)
            worksheet.write(row, col + 5, item[5], text_wrap_format)
            worksheet.write(row, col + 6, item[6])
            worksheet.write(row, col + 7, item[7])
            row += 1            
                
        workbook.close() 
    


def say_goodbye():
    if no_xml == False:
        print("\nAll done! \n\nFind the file(s)in the following directory:\n", directory_path )
    else:
        print("\nLet's try again" )
        get_filelist()
        #TODO add a while loop for better UX 
                

#Let user enter the name of a directory
file_list = get_filelist()

say_goodbye()


