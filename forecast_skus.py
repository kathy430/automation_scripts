# mount google drive
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
import os

forecast_path = 'redacted'

# open forecast path
forecast = gc.open(forecast_path).sheet1

# get data from google sheet
data = forecast.get_all_values()
df = pd.DataFrame(data)

display(df)

# set df columns as first row
df.columns = df.iloc[0]
df = df.iloc[1:]
display(df)

print(df['Brand'].unique())

print(df.columns.tolist())

# only get the months from columns
months = df.columns.tolist()[7:]
print(months)

# change month columns to int
for month in months:
  # Replace empty strings and 'N/A' with 0 before converting to int
  df[month] = df[month].replace('', 0)
  df[month] = df[month].replace('N/A', 0)
  df[month] = df[month].replace('#REF!', 0)
  df[month] = df[month].astype(int)

# formula to expand sets
def expand_sets(row):
  barcode_field = str(row['redacted'])
  expanded = []

  if (barcode_field == 'nan'):
    return pd.DataFrame()

  if '+' in barcode_field:
    barcode_list = barcode_field.split('+')

    for barcode in barcode_list:
      if '*' in barcode:
        qty, barcode = barcode.split('*')
        qty = int(qty)
        barcode = int(barcode)

        # sometimes the barcode is first, not qty so check
        if qty > barcode:
          qty, barcode = barcode, qty

        expanded.append({
            'Barcode': str(barcode),
            'Brand': row['Brand'],
            **{month: row[month] * qty for month in months}
        })
  elif '*' in barcode_field:
    qty, barcode = barcode_field.split('*')
    qty = int(qty)
    barcode = int(barcode)

    # swap qty and barcode if wrong
    if qty > barcode:
      qty, barcode = barcode, qty

    expanded.append({
        'Barcode': str(barcode),
        'Brand': row['Brand'],
        **{month: row[month] * qty for month in months}
    })
  else:
    expanded.append({
        'Barcode': barcode_field,
        'Brand': row['Brand'],
        **{month: row[month] for month in months}
    })

  return pd.DataFrame(expanded)

# expand sets
expanded_list = [expand_sets(row) for _, row in df.iterrows()]
expanded_df = pd.concat(expanded_list, ignore_index=True)

# group by sku
grouped_df = expanded_df.groupby(['Brand', 'Barcode']).sum()

grouped_df.to_csv('redacted')
