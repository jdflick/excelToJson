#! /usr/bin/python3
#^-- change this if you want


import xlrd
from collections import OrderedDict
import json as json
# Open the workbook and select the first worksheet
wb = xlrd.open_workbook('testFile.xlsx')
sh = wb.sheet_by_index(0)
# List to hold dictionaries
attr_list = []

# headers
headers = [str(cell.value) for cell in sh.row(0)]

# Iterate through each row in worksheet and fetch values into dict
for rownum in range(1, sh.nrows):
    attr = OrderedDict()
    row_values = sh.row_values(rownum)
    
    
# get column numbers then pull header data in header name, then append header names in front of dictionary to build arrays       
    for colnum in range(0, sh.ncols):
    
        headerName = headers[colnum]
        attr[headerName] = row_values[colnum]

    attr_list.append(attr)
    
# Serialize the list of dicts to JSON
# it is currently understood that items with JSON escape characters will not import with this code think \ ' " etc.
# this is quick and dirty for simple files
# upcoming versions will solve this

j = json.dumps(attr_list)
# Write to file
with open('data.json', 'w') as f:
    f.write(j)
 

 
 
# this will list all headers, just for testing
# for rownum in range(0, sh.ncols):
    # print(headers[x])
    # x = x +1 