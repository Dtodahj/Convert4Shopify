# Convert4Shopify

A Windows-friendly converter for Lunar Distribution invoices. It accepts original
`.csv`, `.xlsx`, and `.ods` invoice files and creates Shopify product and inventory
imports without requiring Excel, LibreOffice, Python, or an installation process.

## Download

Download [`dist/Shopify_Invoice_Converter.zip`](dist/Shopify_Invoice_Converter.zip),
extract it, and follow the included `README.txt`.

## First-time setup

1. In Shopify Admin, open **Settings → Locations**.
2. Find the active location that should receive the inventory.
3. Double-click **Set Inventory Location.cmd**.
4. Enter the location name exactly as Shopify displays it.

## Convert an invoice

Either drag an original invoice onto **Convert Invoices.cmd**, or place invoices in
the `input` folder and double-click **Convert Invoices.cmd**.

The `output` folder will contain:

- `Shopify_*.csv` — import from **Products → Import**.
- `Inventory_*.csv` — import afterward from **Products → Inventory → Import**.
- `Inventory_*.xlsx` and `Inventory_*.ods` — review/editing copies.
- `Review_*.csv` — validation warnings, excluded rows, and merged duplicates.

Shopify officially imports the CSV copies. Do not open and resave the supplier CSV
before conversion.

## Safety defaults

- Products are created as unpublished drafts.
- Identical duplicate SKUs are merged and their quantities are added.
- Zero-price products are excluded and listed in the review file.
- The converter stops until an exact Shopify inventory location is configured.
- Source invoices are never modified.

For configuration details, see [`README.txt`](README.txt) and [`mapping.json`](mapping.json).

