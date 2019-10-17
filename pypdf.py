"""This module provides access to pdf text extraction and table data extraction.


To run this script on the terminal type:
python pypdf.py '/path/to/filename'

Python Version:
--------------
Version 3.x

Dependencies:
--------------
PyMuPDF


A very very challenging project. :-(
"""
__author__ = 'Okechukwu Peter Diei <okechukwu.diei2@gmail.com>'
__version__ = '0.0.1'
__license__ = 'GNU GPL V3'
__copyright__ = 'Copyright (c) 2019'

import fitz
import time

def pdf_extract(filename, page_no=8):
    """This function performs the extraction of data from a pdf document.
    
    Parameters:
    --------------
    1. filename: the pdf file to be parsed
    2. page_no: the page number on the pdf document that contains the data to be extracted
    """
    # validate that the document page number is the required page number
    if not page_no in (8, 9):
        raise ArgumentError('Function is tailored for just page(s) 8 & 9. Default page number is 8.')
    
    EXTRACTION_TYPE = 'dict'
    output_txt_file = f'output_page_{page_no}.txt'
    output_csv_file = f'output_page_{page_no}.csv'
    xmin = 0.0
    xmax = 0.0
    ymax = 0.0
    ymin = 0.0
    non_table_text = []
    table_text = []
    
    with fitz.open(filename) as document:
        # load page of the pdf and check if the page number is less than 8
        # I did this cos this function is hardcoded to deal with page 8, and 9
        if document.pageCount < 8:
            raise IOError('Number of pages is too small. You need at least 9 pages')
        
        # the fitz module uses a 0-based page numbering system so page 8 
        # is actually referenced by index number 7. So I had to subtract 1 from
        # the page_no passed as an argument
        page = document.loadPage(page_no - 1) 
        page_data = page.getText(EXTRACTION_TYPE)  # extract page as dictionary so we can access page meta data
        
        # page number is page 8
        if page_no == 8:
            for i in page_data['blocks']:
                for j in i['lines']:
                    xmin, ymin, xmax, ymax = j['bbox'] # unpack the bbox tuple
                    
                    if xmin <= 526.6923828125 and ymin <= 216.95242309570312 and xmax <= 544.8453979492188 and ymax <= 209.52511596679688:
                        for k in j['spans']:
                            non_table_text.append(k['text'])
                    elif xmin >= 56.747901916503906 and ymin >= 216.95242309570312 and xmax >= 99.24411010742188 and ymax >= 227.0414276123047:   
                        for k in j['spans']:
                            table_text.append(k['text'])
                            
            # save page text to text file
            with open(output_txt_file, "w") as output_file: 
                text_title = f'Page {page_no} Text'
                output_file.write(text_title + '\n')
                for data in non_table_text:
                    output_file.write(data + '\n')
            
            print(f'{output_txt_file} created.')
            time.sleep(1)
                    
            # save the table data to csv file
            # I had some issues using the csv module so had to work a hack to save the file as csv 
            # I opened the file using a context manager then created two empty lists that would 
            # hold a columns of the data in the table
            with open(output_csv_file, 'w', newline='') as output_file: 
                name_list = []
                amount_list = []
                
                # iterate through the table data to obtain the index that will be used to 
                # determine odd and even columns
                for index, val in enumerate(table_text): 
                    count = index + 1
                    if count % 2 == 0:
                        amount_list.append(val) 
                    else:
                        name_list.append(val)
                
                for key, value in zip(name_list, amount_list):
                    # csv files use ',' to separate fields. The numeric data in the pdf file contains a ','
                    # so I needed to replace this with an empty string to prevent any conflicts
                    # it is just a hack and the comma can be re-inserted if there is a need for it using
                    # f-strings
                    value = int(value.replace(',', ''))
                    key = key.replace(',', '')
                    output_file.write(f'{key}, {value:5d}\n')
                
            print(f'{output_csv_file} created.')
            time.sleep(1)
        
        # page number is page 9
        elif page_no == 9:
            for i in page_data['blocks']:
                for j in i['lines']:
                    xmin, ymin, xmax, ymax = j['bbox'] # unpack the bbox tuple
                    if (xmin == 3.0551986694335938 or xmin >= 45.40998077392578 or  xmin >= 45.409996032714844 or (268.51611328125 <= xmin <= 335.7107849121094)) and (ymin == -28.3740234375 or ymin >= 410.6461181640625 or (167.8187255859375 <= ymin <= 197.818603515625)) and (xmax == 293.8803405761719 or xmax >= 108.6269302368164 or (96.89763641357422 <= xmax <= 359.02099609375)) and (ymax == -15.5537109375 or ymax >= 418.4931335449219 or ymax >= 515.7898559570312 or (178.18673706054688 <= ymax <= 208.18661499023438)):
                        for k in j['spans']:
                            non_table_text.append(k['text'])
                    else:
                        for k in j['spans']:
                            table_text.append(k['text'])
            
            # save page text to text file
            with open(output_txt_file, 'w') as output_file: 
                text_title = f'Page {page_no} Text'
                output_file.write(text_title + '\n')
                for data in non_table_text:
                    output_file.write(data + '\n')
            
            print(f'{output_txt_file} created.')
            time.sleep(1)
            
            # hack to concatenate '*' with the str 'Skills Matter Limited'
            index = table_text.index('*')
            table_text[index - 1] += table_text[index]
            del table_text[index]
        
            # save the table data to csv file
            with open(output_csv_file, 'w', newline='') as output_file: 
                # iterate through the table data to obtain the index that will be used to 
                # determine odd and even columns
                count = 0
                for data in table_text:
                    count += 1
                    # I needed to remove all ',' character in numeric values
                    if ',' in data:
                        data = data.replace(',', '')
                    if count % 6 == 0:
                        output_file.write(f'{data}\n')
                    else:
                        output_file.write(f'{data}, ')
            
            print(f'{output_csv_file} created.')
            time.sleep(1)
    
if __name__ == '__main__':
    import sys
    
    start_time = time.time()
    sleep_time = 4 # time (in secs) that the script was idle by calling the sleep function
    
    if len(sys.argv) < 2:
        raise RuntimeError('No input file name specified')
    
    filename = sys.argv[1]
    
    print('Python script running...') 
    
    pdf_extract(filename) # uses the default number 8
    pdf_extract(filename, 9)
    
    end_time = time.time()
    total_time = end_time - start_time - sleep_time
    
    print('Program executed: ', end='')
    print(f'script ran in exactly {total_time:0.2f} secs and {end_time - start_time:0.2f} secs gross time')
