import pypdf, tabula, re, os, shutil, copy, getopt, sys
from pypdf import generic
from pypdf import Transformation, PageObject
from pypdf.generic import RectangleObject

# Setup number of pages
# Outputs number_of_pages global
def extract_information(pdf_path):
    global number_of_pages

    with open(pdf_path, 'rb') as f:
        pdf = pypdf.PdfReader(f)
        number_of_pages = len(pdf.pages)

# Finds the page numbers for each loop banner e.g Loop 1 Output Points
# Outputs page_numbers and loop_numbers globals
def findPages(pdf_path):
   x1 = 51.690625000000004
   y1 = 50.946875
   x2 = 543.309375
   y2 = 70.284375

   top = min(y1, y2)
   left = min(x1, x2)
   bottom = max(y1, y2)
   right = max(x1, x2)

   area_coordinates = [top, left, bottom, right]

   # Extract tables from the PDF
   tables = tabula.read_pdf(pdf_path, pages='all', area=area_coordinates)

   # Define the pattern to search for
   pattern = r'Loop\s([1-9]|1[0-6])\sOutput\sPoints'

   # Initialize a list to store the page numbers
   global page_numbers
   page_numbers = []

   # Initialise a list to store loop numbers
   global loop_numbers
   loop_numbers = []

   # Iterate over the extracted tables
   for i, table in enumerate(tables, start=1):
      # Convert the table to a string
      table_str = table.to_string()

      # Search for the pattern in the table string
      match = re.search(pattern, table_str)
      if match:
         page_numbers.append(i)
         loop_numbers.append(match.group(1))


def check_temp_folder(outputFolder):
    folder_name = outputFolder
    current_directory = os.getcwd()
    folder_path = os.path.join(current_directory, folder_name)

    if os.path.isdir(folder_path):
        print(f"The folder '{folder_name}' exists in the current directory.")
        # Remove all files within the folder
        for file_name in os.listdir(folder_path):
            file_path = os.path.join(folder_path, file_name)
            os.remove(file_path)
        print(f"All files within '{folder_name}' have been removed.")
    else:
        print(f"The folder '{folder_name}' does not exist in the current directory.")
        # Create the folder
        os.mkdir(folder_path)
        print(f"The folder '{folder_name}' has been created.")

# Extracts the loop files into separate pdf files
# E.g Loop 1 pages in one pdf for editing
# Gets called recursively in another function
def extract_pages(input_path, output_path, start_page, end_page):
    with open(input_path, "rb") as input_file, open(output_path, "wb") as output_file:
        reader = pypdf.PdfReader(input_file)

        # Create new pdf with all pages in    
        writer = pypdf.PdfWriter()

        # number of pages
        numPages = len(range(start_page-1,end_page))

        if numPages % 2 == 1 & numPages > 1:
            numPages = numPages + 1 # add one if odd and greater than 1

        # create blank page to hold all pages in one page
        merged_page = pypdf.PageObject.create_blank_page(None, 595, (741*numPages))        

        y = 0
        i = 1
        for page_number in range(start_page - 1, end_page):
            page = reader.pages[page_number]

            if page_number == start_page - 1:
                # alter crop box of first page to remove header
                RO = RectangleObject((53,50, 542,771))
                page.cropbox = RO
                
            else:
                # alter cropbox to just content of page
                RO = RectangleObject((53,50, 542,791))
                page.cropbox = RO
                

            # if i < numPages:
            merged_page.add_transformation(Transformation().scale(1,1).translate(0, y))

            if i > 2:
                merged_page.add_transformation(Transformation().scale(1,1).translate(0, -18))

            merged_page.merge_page(page, expand=False)


            y = float(page.cropbox.height) + 5
            i += 1
        
        merged_page.add_transformation(Transformation().scale(1,1).translate(0, -50))
        writer.add_page(merged_page)
        writer.write(output_file)

# Splits pages into separate loops
# Ouptuts to temp directory
def split_all_loops(input_path, output_folder):
    pdfName = get_pdf_name(input_path)

    i = 0
    while i < len(page_numbers):
        initialPageNumber = page_numbers[i]
        loopNumber = loop_numbers[i]

        # if i = len
        if i + 1 == len(page_numbers):
            endPageNumber = number_of_pages
        else:
            endPageNumber = page_numbers[i+1] - 1

        collected_tables = [] # list of tables
        output_path = output_folder + pdfName + " Loop " + loopNumber + ".pdf"

        extract_pages(input_path, output_path, initialPageNumber, endPageNumber)
        
        i = i + 1

# Get the basename of the file without extension
def get_pdf_name(input_path):
    file_name = os.path.basename(input_path)
    file_name, _ = os.path.splitext(file_name)
    return str.join('', file_name)

# combine all pages into one
def function_to_path(folder_path, function):
    for filename in os.listdir(folder_path):
        file_path = os.path.join(folder_path, filename)
        if os.path.isfile(file_path) and filename.lower().endswith(".pdf"):
            function(file_path)


# Example function to be performed on each PDF
def join_all_loops(file_path):
    with open(file_path, "rb") as input_file:
        reader = pypdf.PdfReader(input_file)
        # Create new pdf with all pages in    
        writer = pypdf.PdfWriter()  

        # Calculate how tall and wide the page should be
        numPages = len(reader.pages) + 1
        width = reader.pages[0].cropbox.width + 100
        pageHeight = reader.pages[0].cropbox.height

        pageHeight = (pageHeight * numPages)

        # create blank page to hold all pages in one page
        merged_page = pypdf.PageObject.create_blank_page(None, width, pageHeight)
        
        y = 0
        for page in reader.pages:
            merged_page.add_transformation(Transformation().scale(1,1).translate(0, y))
            merged_page.merge_page(page, expand=False)
            
            y += page.mediabox.height

        writer.add_page(merged_page)
        # at end of file
        with open(file_path, "wb") as output_file:
            writer.write(output_file)

def split_pdf_vertically(input_file):
    with open(input_file, 'rb') as file:
        reader = pypdf.PdfReader(file)
        total_pages = len(reader.pages)

        writer = pypdf.PdfWriter()

        split_x = 297

        for page_number in range(total_pages):
            page = reader.pages[page_number]
            page2 = reader.pages[page_number]
            page_width = page.mediabox.width
            page_height = page.mediabox.height

            # Crop the left half of the page
            left_area = RectangleObject((0, 0, split_x, page_height))
            page.cropbox = left_area
            writer.add_page(page)

            # Crop the right half of the page
            right_area = RectangleObject((split_x, 0, page_width, page_height))
            page2.cropbox = right_area
            writer.add_page(page2)

        # Write the output to a new PDF file
        with open(input_file, 'wb') as output:
            writer.write(output)


# Repeat for each file in the input directory




# Functions apply to all in the output folder



# Heading locations
# x1 = 51.690625000000004
# y1 = 50.946875
# x2 = 543.309375
# y2 = 70.284375

# Page with headings
# "x1": 53.550000000000004,
# "x2": 542.19375,
# "y1": 76.978125,
# "y2": 791.7218750000001,

# Alternate pages
# "x1": 53.550000000000004,
# "x2": 542.19375,
# "y1": 50.946875,
# "y2": 791.7218750000001,

if __name__ == '__main__':
    inputFolder = ''
    outputFolder = ''
    opts, args = getopt.getopt(sys.argv[1:],"hi:o:",["ifolder=", "ofolder="])
    for opt, arg in opts:
        if opt == '-h':
            file_name = os.path.basename(__file__)
            script_name = ''.join(os.path.splitext(file_name)) #str.join shorthand ''.join()
            print ( script_name + ' -i <inputFolder> -o <outputFolder>')
            sys.exit()
        elif opt in ("-i", "--ifolder"):
            if arg != "":
                if os.path.isdir(arg):
                    inputFolder = arg
                else:
                    print("Input directory not valid")
                    sys.exit()
            else:
                sys.exit()
        elif opt in ("-o", "--ofolder"):
            if arg != "":
                if os.path.isdir(arg):
                    outputFolder = arg
                else:
                    print("Output directory not valid")
                    sys.exit()
            else:
                sys.exit()         

    # path = 'Node10-ED-Modular-Level 0.pdf'
    print("% Scanning %")

    print(outputFolder)

    # Check the temp folder for files / and it exists
    check_temp_folder(outputFolder)

    for filename in os.listdir(inputFolder):
        if filename.endswith(".pdf"):
            file_path = os.path.join(inputFolder, filename)
            # Get length of pdf
            extract_information(file_path)

            # Find page numbers and loop numbers
            findPages(file_path)

            # Split all loops to individual pdfs
            split_all_loops(file_path, outputFolder)

    # Join all pages in each pdf
    function_to_path(outputFolder, join_all_loops)

    # Split pdfs vertically
    function_to_path(outputFolder, split_pdf_vertically)

    print("% DONE %")