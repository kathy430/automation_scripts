

from google.colab import drive
drive.mount('/content/drive')

from google.colab import auth
auth.authenticate_user()

import gspread
from google.auth import default
creds, _ = default()

gc = gspread.authorize(creds)

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import glob
import os

# path to data folder
data_path = ""

# get a list of all excel files
file_list = glob.glob(os.path.join(data_path, '**/*.xlsx'))

# marker list for plot visualization
markers = ['.', 'o', 'v', '^', '<', '>', '8', 's', 'p', 'P', '*', 'h', 'H', '+', 'x', 'X', 'D', 'd']

# compile all data
all_data = []
for file in file_list:
  # get sheet name
  file_name = file.split('/')[-1]
  sheet_name = file_name.split('.')[0]

  df = pd.read_excel(file, sheet_name=sheet_name)
  month = pd.to_datetime(sheet_name, format='%b %Y')
  df['Month'] = month
  all_data.append(df)

monthly_df = pd.concat(all_data, ignore_index=True)

# only keep the columns that we care about
monthly_df = monthly_df[['(Child) ASIN', 'Brand', 'Barcode', 'Title', 'Units Ordered', 'Ordered Product Sales', 'Month']]


# make entire brand column capitalized
monthly_df['Brand'] = monthly_df['Brand'].str.upper()

# get all different brands from monthly df
brand_list = monthly_df['Brand'].unique().astype(str)

print(np.sort(brand_list))

"""# pick a brand name from the brand list ^ (the brand name needs to match exactly, so just copy and paste it)"""

# put brand name in quotes
chosen_brand = ""

# select timeframe of sales data that you want format is YYYY-MM-DD (all dates should have DD as 01)
start_date = '2026-01-01'
end_date = '2026-01-31'

start_date = pd.to_datetime(start_date)
end_date = pd.to_datetime(end_date)

# get chosen brand df
chosen_brand_df = monthly_df[monthly_df['Brand'] == chosen_brand]
chosen_brand_df = chosen_brand_df[(chosen_brand_df['Month'] >= start_date) & (chosen_brand_df['Month'] <= end_date)]

# combine same asins by month
combined_asin_df = chosen_brand_df.groupby(['Month', '(Child) ASIN'])['Ordered Product Sales'].sum().reset_index()


# create visual data to show brand's monthly asin profit
monthly_sorted_df = combined_asin_df.sort_values(by='Month')

plt.figure(figsize=(50, 20))

# for each ASIN, plot its line on the same chart
for i, asin in enumerate(monthly_sorted_df['(Child) ASIN'].unique()):
  # get asin's monthly profit
  asin_data = monthly_sorted_df[monthly_sorted_df['(Child) ASIN'] == asin]
  plt.plot(
      asin_data['Month'],
      asin_data['Ordered Product Sales'],
      marker=markers[i % len(markers)], # cycle through marker list
      label=asin
  )

plt.title(f'Monthly Profit per ASIN - {chosen_brand}')
plt.xlabel('Month')
plt.ylabel('Profit')
plt.legend(bbox_to_anchor=(1.04, 1), loc='upper left')
plt.show()

"""## ASIN Links for Monthly Profit Plot

Below are the ASINs from the 'Monthly Profit per ASIN' plot, with clickable links to their Amazon product pages:


"""

asin_links_md = ""
for asin in monthly_sorted_df['(Child) ASIN'].unique():
    asin_links_md += f"- [{asin}](https://www.amazon.com/dp/{asin})\n"

# Display the generated markdown content
from IPython.display import display, Markdown
display(Markdown(asin_links_md))

# show graph by sku
# formula to parse through sets
def expand_sets(row):
  barcode_field = str(row['Barcode'])
  expanded = []

  if '+' in barcode_field:
    barcodes = barcode_field.split('+')

    for barcode in barcodes:
      if '*' in barcode:
        qty, sku = barcode.split('*')
        qty = int(qty)
      else:
        sku = barcode
        qty = 1

      if qty == 1:
        expanded.append({
            '(Child) ASIN': row['(Child) ASIN'],
            'Brand': row['Brand'],
            'Barcode': sku,
            'Title': row['Title'],
            'Units Ordered': int(row['Units Ordered']),
            'Ordered Product Sales': float(row['Ordered Product Sales']) / len(barcodes), # split sales evenly amongst items
            'Month': row['Month']
        })
  elif '*' in barcode_field:
    qty, sku = barcode_field.split('*')
    qty = int(qty)

    expanded.append({
        '(Child) ASIN': row['(Child) ASIN'],
        'Brand': row['Brand'],
        'Barcode': sku,
        'Title': row['Title'],
        'Units Ordered': int(row['Units Ordered']) * qty,
        'Ordered Product Sales': float(row['Ordered Product Sales']),
        'Month': row['Month']
    })
  else:
    expanded.append({
        '(Child) ASIN': row['(Child) ASIN'],
        'Brand': row['Brand'],
        'Barcode': row['Barcode'],
        'Title': row['Title'],
        'Units Ordered': int(row['Units Ordered']),
        'Ordered Product Sales': float(row['Ordered Product Sales']),
        'Month': row['Month']
    })
  return pd.DataFrame(expanded)

# expand sets
expanded_list = [expand_sets(row) for _, row in chosen_brand_df.iterrows()]
expanded_df = pd.concat(expanded_list, ignore_index=True)

sku_sorted_df = expanded_df.groupby(['Month', 'Barcode'])['Ordered Product Sales'].sum().reset_index()
total_sku_ordered = expanded_df.groupby('Barcode')['Units Ordered'].sum().reset_index()

# write to google sheet sum of barcode data
gsheet = gc.open('Total Barcode Units Ordered')

# check if worksheet already exists
try:
  new_ws = gsheet.worksheet('Sheet1')
  print('Worksheet already exists. Clearing existing data...')
  new_ws.clear()
except gspread.exceptions.WorksheetNotFound:
  print('Worksheet does not exist. Creating new worksheet...')
  new_ws = gsheet.add_worksheet(title='Sheet1', rows='1', cols='1')

# convert dataframe to a list of lists
data_to_write = [total_sku_ordered.columns.values.tolist()] + total_sku_ordered.values.tolist()

# write data to google sheet
new_ws.update(data_to_write)

print('Data written to Google Sheet.')



plt.figure(figsize=(30, 20))

# for each sku, plot its line on the same chart
for i, sku in enumerate(sku_sorted_df['Barcode'].unique()):
  # skip empty barcodes
  if sku == '#VALUE!':
    continue
  # get sku's monthly profit
  sku_data = sku_sorted_df[sku_sorted_df['Barcode'] == sku]
  plt.plot(
      sku_data['Month'],
      sku_data['Ordered Product Sales'],
      marker=markers[i % len(markers)], # cycle through marker list
      label=sku
  )

plt.title(f'Monthly Profit per SKU - {chosen_brand}')
plt.xlabel('Month')
plt.ylabel('Profit')
plt.legend(bbox_to_anchor=(1.04, 1), loc='upper left')
plt.show()

