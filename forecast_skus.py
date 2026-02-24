import pandas as pd
import numpy as np
import os

forecast_path = ""

# open excel file
forecast = pd.read_excel(forecast_path)

print(forecast.columns)

# formula to expand sets
def expand_sets(row):
  barcode_field = str(row['Barcode'])
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
            #'6M': row['6M'] * qty,
            'Sum of 25-Oct': row[pd.Timestamp('2026-10-25 00:00:00')] * qty,
            'Sum of 25-Nov': row[pd.Timestamp('2026-11-25 00:00:00')] * qty,
            'Sum of 25-Dec': row[pd.Timestamp('2026-12-25 00:00:00')] * qty,
            'Sum of 26-Jan': row[pd.Timestamp('2026-01-26 00:00:00')] * qty,
            'Sum of 26-Feb': row[pd.Timestamp('2026-02-26 00:00:00')] * qty,
            'Sum of 26-Mar': row[pd.Timestamp('2026-03-26 00:00:00')] * qty,
            'Sum of 26-Apr': row[pd.Timestamp('2026-04-26 00:00:00')] * qty,
            'Sum of 26-May': row[pd.Timestamp('2026-05-26 00:00:00')] * qty,
            'Sum of 26-Jun': row[pd.Timestamp('2026-06-26 00:00:00')] * qty,
            'Sum of 26-Jul': row[pd.Timestamp('2026-07-26 00:00:00')] * qty,
            'Sum of 26-Aug': row[pd.Timestamp('2026-08-26 00:00:00')] * qty,
            'Sum of 26-Sep': row[pd.Timestamp('2026-09-26 00:00:00')] * qty
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
        #'6M': row['6M'] * qty,
        'Sum of 25-Oct': row[pd.Timestamp('2026-10-25 00:00:00')] * qty,
        'Sum of 25-Nov': row[pd.Timestamp('2026-11-25 00:00:00')] * qty,
        'Sum of 25-Dec': row[pd.Timestamp('2026-12-25 00:00:00')] * qty,
        'Sum of 26-Jan': row[pd.Timestamp('2026-01-26 00:00:00')] * qty,
        'Sum of 26-Feb': row[pd.Timestamp('2026-02-26 00:00:00')] * qty,
        'Sum of 26-Mar': row[pd.Timestamp('2026-03-26 00:00:00')] * qty,
        'Sum of 26-Apr': row[pd.Timestamp('2026-04-26 00:00:00')] * qty,
        'Sum of 26-May': row[pd.Timestamp('2026-05-26 00:00:00')] * qty,
        'Sum of 26-Jun': row[pd.Timestamp('2026-06-26 00:00:00')] * qty,
        'Sum of 26-Jul': row[pd.Timestamp('2026-07-26 00:00:00')] * qty,
        'Sum of 26-Aug': row[pd.Timestamp('2026-08-26 00:00:00')] * qty,
        'Sum of 26-Sep': row[pd.Timestamp('2026-09-26 00:00:00')] * qty
    })
  else:
    expanded.append({
        'Barcode': barcode_field,
        #'6M': row['6M'],
        'Sum of 25-Oct': row[pd.Timestamp('2026-10-25 00:00:00')],
        'Sum of 25-Nov': row[pd.Timestamp('2026-11-25 00:00:00')],
        'Sum of 25-Dec': row[pd.Timestamp('2026-12-25 00:00:00')],
        'Sum of 26-Jan': row[pd.Timestamp('2026-01-26 00:00:00')],
        'Sum of 26-Feb': row[pd.Timestamp('2026-02-26 00:00:00')],
        'Sum of 26-Mar': row[pd.Timestamp('2026-03-26 00:00:00')],
        'Sum of 26-Apr': row[pd.Timestamp('2026-04-26 00:00:00')],
        'Sum of 26-May': row[pd.Timestamp('2026-05-26 00:00:00')],
        'Sum of 26-Jun': row[pd.Timestamp('2026-06-26 00:00:00')],
        'Sum of 26-Jul': row[pd.Timestamp('2026-07-26 00:00:00')],
        'Sum of 26-Aug': row[pd.Timestamp('2026-08-26 00:00:00')],
        'Sum of 26-Sep': row[pd.Timestamp('2026-09-26 00:00:00')]
    })

  return pd.DataFrame(expanded)

# expand sets
expanded_list = [expand_sets(row) for _, row in forecast.iterrows()]
expanded_df = pd.concat(expanded_list, ignore_index=True)

# group by sku
grouped_df = expanded_df.groupby(['Barcode']).sum()
#print(grouped_df)

# get brand and item description
item_df = forecast[['Barcode', 'Brand', 'Description']].copy()
item_df['Barcode'] = item_df['Barcode'].astype(str)

final_df = pd.merge(grouped_df, item_df, on='Barcode', how='left')
final_df = final_df[['Barcode', 'Brand', 'Description', 'Sum of 25-Oct', 'Sum of 25-Nov', 'Sum of 25-Dec',
                    'Sum of 26-Jan', 'Sum of 26-Feb', 'Sum of 26-Mar', 'Sum of 26-Apr', 'Sum of 26-May',
                    'Sum of 26-Jun', 'Sum of 26-Jul', 'Sum of 26-Aug', 'Sum of 26-Sep']]


sku_forecast_path = ""

sorted_brands = np.sort(final_df['Brand'].astype(str).unique())

# create excel sheet if it doesn't exist
if not os.path.exists(sku_forecast_path):
  df_empty = pd.DataFrame()
  df_empty.to_excel(sku_forecast_path)

# write to excel
with pd.ExcelWriter(sku_forecast_path, mode='a', if_sheet_exists='replace') as writer:
  for brand in sorted_brands:
    # remove invalid character for sheet name
    if '/' in str(brand):
      brand_name = brand.split('/')[0]
    else:
      brand_name = brand

    # skip nan brands
    if brand_name == 'nan':
      continue

    # get specific brand dataframe
    brand_df = final_df[final_df['Brand'] == brand]

    # write to excel
    brand_df.to_excel(writer, sheet_name=f'{brand_name} FCST', index=False)

