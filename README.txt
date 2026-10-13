SHOPIFY INVOICE CONVERTER - VERSION 1.2

WHAT IT DOES
Converts original Lunar Distribution .csv, .xlsx, or .ods invoices into:
  1. Shopify_*.csv - the product import file.
  2. Inventory_*.csv - inventory quantities for your store location.
  3. Review_*.csv - warnings, excluded rows, and merged duplicates.
  4. Inventory_*.xlsx and Inventory_*.ods - spreadsheet copies matching the
     Shopify inventory template for review or editing.

No Excel, LibreOffice, Python, or installation is required.

FIRST-TIME SETUP (REQUIRED)
1. In Shopify Admin, open Settings > Locations.
2. Find the active location that should receive this inventory.
3. Double-click "Set Inventory Location.cmd".
4. Type or paste the location name exactly as Shopify displays it, then press Enter.

Repeat setup only if the Shopify location name changes. The converter stops with a
clear message if the location has not been configured.

CONVERT AN INVOICE
Easy method:
  Drag an original .csv, .xlsx, or .ods invoice directly onto
  "Convert Invoices.cmd". Do not open or resave the supplier CSV first.

Folder method:
  1. Put one or more invoices in the input folder.
  2. Double-click "Convert Invoices.cmd".

Finished files appear in the output folder. Source invoices are never changed.

IMPORT INTO SHOPIFY
1. Open Review_*.csv and check any REVIEW or EXCLUDED entries.
2. In Shopify Admin, go to Products > Import and select Shopify_*.csv.
3. After the product import finishes, go to Products > Inventory > Import and
   select Inventory_*.csv.
4. Check Shopify's import summary and confirm the total inventory quantity.

The inventory file is important. Shopify can ignore inventory quantities in a
product import when the store has multiple or app-managed locations.

Shopify officially imports CSV files. The XLSX and ODS inventory copies are
provided to view or edit the same data, but use Inventory_*.csv for Shopify.

DEFAULT BEHAVIOR
- Products are created as DRAFT and are not published to the online store.
- Retail becomes Price; Discounted Price becomes Cost per item.
- Qty becomes inventory at the configured Shopify location.
- Code becomes SKU; UPC becomes Barcode.
- Identical duplicate SKUs are combined and their quantities are added.
- Zero-price products are excluded and listed in the review file.
- The product CSV follows the supplied 57-column Shopify template.

CONFIGURATION
Advanced users can open mapping.json in Notepad to change vendor, status, tags,
category mappings, publishing defaults, source column names, or zero-price handling.
For normal use, configure the location with "Set Inventory Location.cmd" instead
of editing mapping.json manually.

SYSTEM REQUIREMENT
Windows PowerShell 5.1 or newer (included with supported Windows versions).
