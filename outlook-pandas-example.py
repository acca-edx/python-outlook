### INFO
# This file only contains comments relevant to the new functions called.
#   see previous example if you require more information.

import win32com.client

### Read in an excel spreadsheet with pandas
# Import pandas and give it label pd (for simplicity)
# Using pandas built in function read_excel we specify the file location
#   and a sheet_name with in that. 
import pandas as pd
excel_data = pd.read_excel(r'C:\Users\<name>\Documents\excel_example.xlsx',
                           sheet_name="Sheet1") 

### Create an outlook Email as per previous example
o = win32com.client.Dispatch("Outlook.Application")

### Loop through the content of the spreadsheet 
# for loop will pass each line of the excel spreads row index.
#   The variable 'index' will be the current line it is processing. 
# Note: Python uses code indents to define a block of code, the indent
#   must be of the same size. (either a tab or number of spaces.)
for index in excel_data.index:
  # Note the 2 spaces before this block of code
  msg = o.CreateItem(0)
  msg.to = excel_data['name'][index]
  msg.Subject = excel_data['subject'][index]
  msg.Body = excel_data['message'][index]
  msg.Attachments.Add(excel_data['attachment'][index])
  msg.Send()
  # This is the end of the for loop code block.

# Anything after here will continue once the for loop block index is
#   complete.
