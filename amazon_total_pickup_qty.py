import pandas as pd

order_file_path = r""
output_file_path = r""


 
data = pd.read_excel(order_file_path)

# function to expand sets into individual SKUs
def expand_sets(row):
  barcode_field = str(row['ORIGINAL JAN']).strip()
  expanded = []
  
  if '\n' in barcode_field:
    barcodes = barcode_field.split('\n')

    for barcode in barcodes:
      expanded.append({
          'Original Jan': barcode.strip(),
          'Total Qty': int(row['Ship out\nQty'])
      })
  elif ' ' in barcode_field:
    barcodes = barcode_field.split(' ')

    for barcode in barcodes:
      expanded.append({
          'Original Jan': barcode.strip(),
          'Total Qty': int(row['Ship out\nQty'])
      })
  else:
    expanded.append({
        'Original Jan': barcode_field.strip(),
        'Total Qty': int(row['Ship out\nQty'])
    })
  
  return pd.DataFrame(expanded)

expanded_list = [expand_sets(row) for _, row in data.iterrows()]
expanded_df = pd.concat(expanded_list, ignore_index=True)

# get total of each SKU
total_qty_df = expanded_df.groupby('Original Jan')['Total Qty'].sum().reset_index()

# write to excel file
total_qty_df.to_excel(output_file_path, index=False)