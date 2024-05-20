# extract_doc_info.py

import sys, getopt, os, re, tabula
import pandas as pd
from PyPDF2 import PdfReader

def extract_information(pdf_path):
    with open(pdf_path, 'rb') as f:
        pdf = PdfReader(f)
        information = pdf.metadata
        number_of_pages = len(pdf.pages)

    txt = f"""
    Information about {pdf_path}: 

    Author: {information.author}
    Creator: {information.creator}
    Producer: {information.producer}
    Subject: {information.subject}
    Title: {information.title}
    Number of pages: {number_of_pages}
    """

    print(txt)
    return information

def extract_table(pdf_path):
   page_numbers = findPages(pdf_path)
   data_pages = page_numbers[:-1] # remove last page number
   lastPage = page_numbers[-1] # last element last page for range

   data_tables = [] # list of all table lists

   # create page ranges
   i = 0
   while i < len(data_pages):
      initialPageNumber = data_pages[i]

      # if i = len
      if i + 1 == len(data_pages):
         endPageNumber = lastPage
      else:
         endPageNumber = data_pages[i+1]

      collected_tables = [] # list of tables

      # read area pdf
      firstPage = read_page_with_area(pdf_path, initialPageNumber)
      collected_tables.append(firstPage)

      # in new loop read all subsequent pages normally and add to list
      for y in range(initialPageNumber + 1, endPageNumber):

         if initialPageNumber + 1 == endPageNumber:
            continue

         table = read_page_only(pdf_path, y)
         collected_tables.append(table)
         
      combined_df = pd.concat([pd.DataFrame(table[0]) for table in collected_tables])

      data_tables.append(combined_df)
      
      i = i + 1
   
   return data_tables
   
def read_page_loop_number(pdf_path, inputPages):
   x1 = 51.690625000000004
   y1 = 50.946875
   x2 = 543.309375
   y2 = 70.284375

   top = min(y1, y2)
   left = min(x1, x2)
   bottom = max(y1, y2)
   right = max(x1, x2)

   area_coordinates = [top, left, bottom, right]
   
   table = tabula.read_pdf(pdf_path, pages=inputPages[0], area=area_coordinates, guess=False)
   
   output_string = ""
   for i, t in enumerate(table):
      output_string += t.to_csv(None, index=False)  # Get CSV content as a string

   regex = r"Loop\s([1-9]|1[0-6])\sBrief"

   match = re.search(regex, output_string)
   if match:
      number = match.group(1)
   
   return number

def read_page_with_area(pdf_path, page_number):
   x1 = 55.40
   y1 = 76.97
   x2 = 529.17
   y2 = 775.35

   top = min(y1, y2)
   left = min(x1, x2)
   bottom = max(y1, y2)
   right = max(x1, x2)

   area_coordinates = [top, left, bottom, right]
   
   table = tabula.read_pdf(pdf_path, pages=page_number, area=area_coordinates, guess=False)

   return table

def read_page_only(pdf_path, page_number):
   x1 = 55.00031249999999
   y1 = 50.203125
   x2 = 539.9253125
   y2 = 767.178125

   top = min(y1, y2)
   left = min(x1, x2)
   bottom = max(y1, y2)
   right = max(x1, x2)

   area_coordinates = [top, left, bottom, right]

   table = tabula.read_pdf(pdf_path, pages=page_number, area=area_coordinates)
   SingleTable = table[0]

   FirstRow = SingleTable.columns[:] # get first row data
   SingleTable.columns = ['Address', 'Type', 'Zone', 'Zone Text', 'Location'] # set the columns
   new_row = []

   for i in FirstRow:
      new_row.append(i)

   SingleTable.loc[-1] = new_row
   SingleTable.index = SingleTable.index + 1
   SingleTable.sort_index(inplace=True)
  
   return table

def findPages(pdf_path):
   x1 = 55.00031249999999
   y1 = 50.203125
   x2 = 539.9253125
   y2 = 767.178125

   top = min(y1, y2)
   left = min(x1, x2)
   bottom = max(y1, y2)
   right = max(x1, x2)

   area_coordinates = [top, left, bottom, right]

   # Extract tables from the PDF
   tables = tabula.read_pdf(pdf_path, pages='all', area=area_coordinates)

   # Define the pattern to search for
   pattern = r'Loop \d+ Brief Points Description'

   # Initialize a list to store the page numbers
   page_numbers = []

   # Iterate over the extracted tables
   for i, table in enumerate(tables, start=1):
      # Convert the table to a string
      table_str = table.to_string()

      # Search for the pattern in the table string
      if re.search(pattern, table_str):
         page_numbers.append(i)

   # Print the page numbers where the pattern was found

   page_numbers.append(findEndPage(tables))

   # print("Page numbers with 'Loop n Brief Points Description' last item is Display Card Brief Points:")
   # print(page_numbers)

   return page_numbers

def findEndPage(tables):
   pattern = r'Display Card Brief Points Description'
   page_number = ''

   for i, table in enumerate(tables, start=1):
      table_str = table.to_string()
      
      if re.search(pattern, table_str):
         page_number = i
         continue

   return page_number

if __name__ == '__main__':
   inputfile = ''
   outputfile = ''
   opts, args = getopt.getopt(sys.argv[1:],"h:i:o:",["ifile=", "ofile="])
   for opt, arg in opts:
      if opt == '-h':
         file_name = os.path.basename(__file__)
         script_name = ''.join(os.path.splitext(file_name)) #str.join shorthand ''.join()
         print ( script_name + ' -i <inputfile>')
         sys.exit()
      elif opt in ("-i", "--ifile"):
         inputfile = arg
      elif opt in ("-o", "--ofile"):
         outputfile = arg

   # path = 'Node10-ED-Modular-Level 0.pdf'
   print("% Scanning %")

   extract_information(inputfile)
   tables_collection = extract_table(inputfile)

   pages = findPages(inputfile)
   LoopOffset = int(read_page_loop_number(inputfile, pages))
   
   for i, table in enumerate(tables_collection):
      table.to_csv(f"{inputfile}_Loop_{i+LoopOffset}.csv", index=False)  # Specify the desired output filename
      print(f"{inputfile}_Loop_{i+LoopOffset}.csv" + " created!")

   print("% DONE %")