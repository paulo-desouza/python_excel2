# objective for this code: better use of functions. 
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
#from openpyxl.styles import Font, PatternFill, Border, Side
from datetime import datetime
import pyexcel
import os


def convert_xls(xls_file):
    """
    xls_file : xls file.
    
    ------------
        takes an XLS file. 
    ------------
        creates a converted XLSX file.
    
    """
    
    pyexcel.save_book_as(file_name=xls_file,
                   dest_file_name=xls_file+"x")
    
    
def delete_file(file):
    """
    file : any file.
    
    ----------
        takes a file.
    ----------
        deletes it. 
    
    """
    
    os.remove(file)
    
    
def scrape_table(worksheet):
    """
    Parameters
    ----------
    worksheet : openpyXL worksheet object.

    Returns
    -------
    List with the table contents.

    """
    # Row and column counter
    row_count = 0
    
    for row in worksheet:
        if not all([cell.value is None for cell in row]):
            row_count += 1
    
    
    # Loop through the rows and columns, appending the information to a list.
    scraped_info = {}
    
    for r in range(1, row_count+1):
        
        for col in range(1, 5):
            
            char = get_column_letter(col)
            
            if char == "A":
                scraped_info[ws[char + str(r)].value] = []
            else:
                scraped_info[ws["A" + str(r)].value] += [ws[char + str(r)].value]         
            
    return scraped_info
    
    
def write_table(client, worksheet, col_start = 1):
    """
    Parameters
    ----------
    table : list with scraped data.
        
    worksheet : worksheet object to write new table onto.

    Returns
    -------
    None.
    """
    # write the tables.
    
    
    
    for i, key in enumerate(client):
        char = get_column_letter(col_start)
        ws[char+str(i+1)].value = key
        
        for j, value in enumerate(client[key]):
            char = get_column_letter(col_start+j+1)
            
            if char == "A": 
                pass
            else:
                ws[char+str(i+1)].value = client[key][j]
            
    
    
            
    
def date_conversion(date_string):
    """
    Parameters
    ----------
    date_string : String containing a date. example:
        "110223"  or  "11/02/21"  or  "blahblahbvlah 110423"

    Returns
    -------
    DateTime Object.
    """
    # strip the symbols in between the dates if existant. 
    
    rawdate = date_string.split(" ")[1]
    y = int('20'+rawdate[4:])
    m = int(rawdate[0:2])
    d = int(rawdate[2:4])
    
    return datetime(y, m, d)
    
    
def get_file_names():
    """
    Returns
    -------
    All filenames of this python file's current directory, excluding itself. 
    
    """
    
    current_dir = os.getcwd()
    
    file_list = os.listdir(current_dir)
    
    file_list.remove("performance.py")
    
    return file_list


def sort_by_date(file_list):
    """
    Parameters
    ----------
    file_list : List of filenames with dates on them.

    Returns
    -------
    Sorted List By Date.
    """
    
    sorted_files = []
    sorted_list = []
    
    for file in file_list:
        sorted_list.append(date_conversion(file))
    
    sorted_list.sort()
    
    for dt in sorted_list:
        for file in file_list:
        
            if date_conversion(file) == dt:
                
                sorted_files.append(file)
    
    return sorted_files


# CODE START

# convert all files to xls

file_list = get_file_names()

for file in file_list:
    if "xlsx" in file:
        pass
    else:
        convert_xls(file)
        delete_file(file)

# get list with converted files
    
file_list = get_file_names()


sorted_datetimes = sort_by_date(file_list)


# scrape all the tables onto this list:

tables = []

for i, file in enumerate(sorted_datetimes):
    wb = load_workbook(file)
    ws = wb.active
    
    tables.append(scrape_table(ws))
        
    
wb = Workbook()
ws = wb.active
ws.title = "Timeline"


col = 1
for list in tables: 
    write_table(list, ws, col)
    col += 6
    
rate_per_child = [381,      #gb
                  356,      #om
                  526,      #tjb
                  394,      #ash
                  370,      #laurel
                  426,      #rock
                  328,      #montg
                  314,      #fm
                  306,      #henri
                  425,      #reston
                  424,      #elk
                  353,      #warr
                  331,      #nott
                  339,      #EN
                  440,      #ALX
                  436,      #canton
                  381]      #melford


capacities = [150,      #gb
              135,      #om
              152,      #tjb
              126,      #ash
              144,      #laurel
              190,      #rock
              147,      #montg
              168,      #fm
              172,      #henri
              172,      #reston
              141,      #elk
              161,      #warr
              141,      #nott
              145,      #EN
              190,      #ALX
              152,      #canton
              150]      #melford

# ^^^       missing BELLONA and COLUMBIA       ^^^
wb.create_sheet("Growth Report")
ws = wb["Growth Report"]
    

titles = []

for i, key in enumerate(tables[4]):
    if i < 3:
        pass
    else:
        titles.append(key)
    
# insert any headers into the titles list
    
for i, col in enumerate(titles):
    ws["A" + str(i+1)].value = col
    
conter = 0
for row in range(2,8):
    for col in range(1, len(tables[len(tables)])):
        if row == 2:
            pass
  

wb.save("TEST1.xlsx")


