"# OCR-Invoice-extracting-" 
OCR Invoice Extraction

A production-oriented OCR pipeline for extracting structured data from Indian e-commerce invoices (Amazon, Blinkit, Flipkart) and exporting them into a standardized Excel template.

Features

OCR-based invoice parsing using EasyOCR

Handles multi-column invoice layouts (2-column, 3-column, asymmetric splits)

Platform-specific logic for:

Amazon

Blinkit

Flipkart

Structured extraction into a common Excel template

Invoice_Header sheet

Table_1 (line items + totals)

Robust regex-based field extraction

Noise handling for OCR artifacts (misread digits, layout shifts)

Supported Fields
Invoice Header

Billing Address

Shipping Address

Invoice Type

Order Number

Invoice Number

Order Date

Invoice Date

Seller Name & Address

Seller PAN / GST

FSSAI License (where available)

Place of Supply / Delivery

Reverse Charge

Amount in Words

Total Tax

Total Amount

Invoice Table

Sl.No

Description

Unit Price

Discount

Quantity

Net Amount

Tax Rate

Tax Type

Tax Amount

Total Amount

Project Structure
T4/
├── Input/
│   └── images/
├── Outputs/
│   ├── Amazon/
│   ├── Blinkit/
│   └── Flipkart/
├── Output Template.xlsx
├── amazon.py
├── blinkit.py
├── flipkart.py
└── README.md
Tech Stack

Python 3.9+

EasyOCR

OpenCV

Pandas

OpenPyXL

Regex-based layout parsing

How It Works

Invoice image is read using EasyOCR

Layout is split dynamically (2 or 3 columns depending on platform)

Header and table regions are processed independently

Platform-specific heuristics normalize extracted values

Cleaned data is written into a standardized Excel template

Notes

Designed for real-world noisy OCR output

Easily extensible to new vendors

Optimized for accounting / ERP ingestion workflows

Author

Hemanth Padala