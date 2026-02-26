"""
How to use

1. Paste PDF file path after "pdf_path = " on line 12
2. Make sure there is a r before the path in quotes
3. Press play at the top right
"""

import pdfplumber
import pandas as pd
import openpyxl

pdf_path = r""
output_path = r""

pdf = pdfplumber.open(pdf_path)

# create own settings to get tables
table_settings = {
  'vertical_strategy': 'explicit',
  'horizontal_strategy': 'lines',
  'explicit_vertical_lines': [85, 195, 243, 333, 360, 430, 542, 545, 590, 680, 707],
}

# function to clean up rows
def clean_rows(row):
  new_row = []

  for item in row: # get each string in row
    if (item == None or item == ''): # skip useless rows
      return new_row

    new_string = ''
    string_len = len(item)

    for i, char in enumerate(item): # iterate over each char in the string
      if char == '\n':
        if i < string_len - 1 and (item[i+1].islower() or item[i-1] == '-'):
          continue   # just skip
        else:
          new_string = new_string + ' '   # replace \n with a space only if next char is start of a new word
      else:
        new_string = new_string + char    # no change, just add the char

    new_row.append(new_string) # add modified string

  return new_row

COLUMNS = ['Product Name', 'SKU', 'Seller SKU', 'Qty']
order_data = []

# read pdf
with pdfplumber.open(pdf_path) as pdf:
  for page in pdf.pages:
    tables = page.extract_tables(table_settings)

    for table in tables:
      for row in table: # clean up table
        clean_row = clean_rows(row)
        if clean_row != []:
          #print(page, clean_row)
          order_data.append(clean_row)

# convert order data into a dataframe
order_df = pd.DataFrame(order_data)
order_df.columns = COLUMNS

# convert qty column to int
order_df['Qty'] = order_df['Qty'].astype(int)

# group by sku to get total quantities
grouped_df = order_df.groupby(['Seller SKU', 'SKU'], as_index=False)['Qty'].sum()
grouped_df.to_excel(output_path, index=False)

# adjust column widths in excel
wb = openpyxl.load_workbook(output_path)
ws = wb['Sheet1']

for col in ws.columns:
  # only qty column is smaller
  if col[0].column_letter == 'C':
    ws.column_dimensions[col[0].column_letter].width = 10
  else:
    ws.column_dimensions[col[0].column_letter].width = 40

wb.save(output_path)
