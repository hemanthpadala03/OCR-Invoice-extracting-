# OCR Invoice Extraction Pipeline

A robust OCR-based system to extract **structured accounting data** from Indian e-commerce invoices (Amazon, Flipkart, Blinkit) and export them into a **standardized Excel template** suitable for ERP, accounting, and reconciliation workflows.

This project is designed to handle **real-world OCR noise**, **multi-column invoice layouts**, and **platform-specific invoice formats**.

---

## 🚀 Key Features

* OCR using **EasyOCR**
* Supports **Amazon, Flipkart, Blinkit** invoices
* Handles **2-column and 3-column layouts**
* Platform-specific parsing logic
* Extracts:

  * Invoice Header data
  * Line-item table data
* Outputs data into a **common Excel template**
* Built for **accounting & ERP ingestion**

---

## 📄 Extracted Data

### Invoice Header (`Invoice_Header` sheet)

* Billing Address
* Shipping Address
* Invoice Type
* Order Number
* Invoice Number
* Order Date
* Invoice Date
* Seller Name
* Seller Address
* Seller PAN
* Seller GST
* FSSAI License (if available)
* Place of Supply
* Place of Delivery
* Reverse Charge
* Amount in Words
* Total Tax
* Total Amount

### Invoice Table (`Table_1` sheet)

* Sl.No
* Description
* Unit Price
* Discount
* Quantity
* Net Amount
* Tax Rate
* Tax Type (IGST / CGST / SGST)
* Tax Amount
* Total Amount

---

## 🧠 Supported Platforms

| Platform | Layout Type                     | Status     |
| -------- | ------------------------------- | ---------- |
| Amazon   | 2-column                        | ✅ Complete |
| Blinkit  | Asymmetric 2-column             | ✅ Complete |
| Flipkart | 3-column (0–25 / 25–50 / 50–75) | ✅ Complete |

---

## 📁 Project Structure

```
T4/
├── Input/
│   └── images/              # Invoice images
├── Outputs/
│   ├── Amazon/
│   ├── Blinkit/
│   └── Flipkart/
├── Output Template.xlsx     # Standard Excel template
├── amazon.py                # Amazon invoice extractor
├── blinkit.py               # Blinkit invoice extractor
├── flipkart.py              # Flipkart invoice extractor
├── README.md
```

---

## ⚙️ Tech Stack

* Python 3.9+
* EasyOCR
* OpenCV
* Pandas
* OpenPyXL
* Regex-based parsing

---

## ▶️ How to Run

### 1. Install dependencies

```bash
pip install easyocr opencv-python pandas openpyxl
```

### 2. Place invoice image

```
Input/images/<invoice_image>.png
```

### 3. Run extractor

```bash
python amazon.py
python blinkit.py
python flipkart.py
```

### 4. Output

* Extracted Excel file is saved inside:

```
Outputs/<Platform>/Output.xlsx
```

---

## 🧪 Design Highlights

* Column-aware OCR (dynamic image slicing)
* OCR noise correction (misread digits, broken tokens)
* Regex generalized to survive layout shifts
* Safe Excel writing without template corruption
* Platform-aware tax handling (IGST vs CGST/SGST)

---

## ⚠️ Notes

* Designed for **real invoices**, not synthetic PDFs
* EasyOCR runs on CPU by default (GPU optional)
* Extendable to new vendors by adding new scripts

---

## 👨‍💻 Author

**Hemanth Padala**
AI / OCR / Document Intelligence

---

## 📌 Future Enhancements

* Batch invoice processing
* PDF support
* Layout auto-detection
* REST API wrapper
* Multi-line item invoices

---

If you want, I can next:

* Add screenshots to README
* Add architecture diagram
* Convert this into a reusable Python package
* Create a demo notebook

Just tell me.
