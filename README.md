# automation_scripts
A repository consisting of the automation scripts I created for work.

Below is a description of all the automation scripts I have uploaded:

# amazon_total_pickup_qty.ipynb
This notebook takes data from weekly ASIN orders for Amazon and generates a list of each SKU and its quantity that the warehouse needs to prepare.

# forecast_skus.ipynb
Because we have ASINs that are a set of items, this notebook breaks up the barcodes in each ASIN and outputs the total monthly forecast for each SKU.

# monthly_asin_data_gen.ipynb
This notebook organizes Amazon sales data to show monthly profit by ASIN and SKU for specified Brands. 

# tiktok_picking_list_gen.ipynb
This notebook takes a TikTok shop packing list pdf file and parses through each page to generate a picking list with total quantities of each item.
