# Import Shopify Data into Google Sheets 

Shopify's rest API makes it difficult to regularly track and analyze data without first tabulating. Google sheets works great as the middle stage of an ETL pipeline for Shopify data. There are Shopify Apps and Google Apps that do this for you, but most of them are paid and require a subscription. 

This is a Google Apps Script that does the heavy lifting to pull data from the shopify Orders API and tabulate in easily accessible google sheets. There is a 5 million cell maximum on google sheets which roughly translates to 100k orders with the default columns. If you have over 100k orders to analyze, this isn't the best solution. This Docker based SQL Server solution may be more appropriate - [pyshopify](https://github.com/webdjoe/pyshopify) 

## Spreadsheet Structure

The following sheets are required to pull data and the headers listed are the default headers. Any Order object key can be used as a header, the script will automatically pull and load the data based on the column headers. The full Shopify Orders object/keys can be found [here](https://shopify.dev/api/admin-rest/2022-04/resources/order#resource-object) 

| Sheet Name            | Default Columns/Keys             |
|-----------------------|----------------------------------|
| orders                | id, number, order_name, name, current_total_discounts, current_total_price, current_subtotal_price, current_total_tax, financial_status, landing_site, created_at, processed_at, processing_method, subtotal_price, source_url, tags, token, total_discounts, total_line_items_price, total_price, total_tax, updated_at|
| shipping_lines        | order_id, ordered_at, code, price, title, source, discounted_price, carrier_identifier|
| discount_applications | order_id, created_at, type, title, value, value_type, code, description, target_type, target_selection, allocation_method, ordered_at|
| refunds               | id, order_id, ordered_at, created_at, note, transaction_id, kind, amount|
| line_items            | id, order_id, sku, name, price, title, vendor, quantity, product_id, variant_id, variant_title, total_discount, fulfillment_status, ordered_at|
| customer              | order_id, ordered_at, id, email, phone, last_name, created_at, first_name, admin_graphql_api_id, updated_at, orders_count, last_order_id|



## How to Use

Copy and paste `shopify_api.gs` and `shopify_parse.gs` into a Google Apps Script file inside a spreadsheet with the required sheet names/columns.

The function  `RunOrders()` will run the main function. It can accept no arguments, where it will default to pull data up to 30 days ago. 

Or it can accept a start and end date in "yyyy-MM-dd" format to pull data between two dates.

```javascript
// Pulls last 30 days of data
RunOrders();

// Get data between Jan 1, 2022 and Jan 31, 2022
RunOrders("2022-01-01", "2022-01-31");

```

