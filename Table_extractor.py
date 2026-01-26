import cv2
import easyocr
import numpy as np
import pandas as pd
from openpyxl import Workbook

# =====================================================
# CONFIG
# =====================================================
IMAGE_PATH = r"C:\Drive_d\Python\F-AI\T4\Input\images\ForwardInvoice_ORD46618691136_page_2.png"
OUTPUT_EXCEL = "Invoice_Table_Cellwise.xlsx"

# =====================================================
# LOAD IMAGE
# =====================================================
img = cv2.imread(IMAGE_PATH)
h, w, _ = img.shape

reader = easyocr.Reader(['en'], gpu=False)

# =====================================================
# STEP 1: FIND "STATE" → CUT BELOW
# =====================================================
ocr_full = reader.readtext(img, detail=1)

state_y_max = None
for box, text, conf in ocr_full:
    if text.strip().lower() == "state":
        state_y_max = int(max(p[1] for p in box))
        break

if state_y_max is None:
    raise RuntimeError("State not detected")

after_state = img[state_y_max + 10 : h, 0:w]

# =====================================================
# STEP 2: FIND "AMOUNT IN" → CUT ABOVE
# =====================================================
ocr_after_state = reader.readtext(after_state, detail=1)

amount_y_min = None
for box, text, conf in ocr_after_state:
    if "amount in" in text.lower():
        amount_y_min = int(min(p[1] for p in box))
        break

if amount_y_min is None:
    raise RuntimeError("Amount in not detected")

table_img = after_state[0 : amount_y_min - 10, 0:w]
th, tw, _ = table_img.shape

# =====================================================
# STEP 3: BINARIZE
# =====================================================
gray = cv2.cvtColor(table_img, cv2.COLOR_BGR2GRAY)
_, bw = cv2.threshold(
    gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU
)

# =====================================================
# STEP 4: DETECT HORIZONTAL LINES
# =====================================================
h_kernel = cv2.getStructuringElement(
    cv2.MORPH_RECT, (int(tw * 0.4), 1)
)
h_lines = cv2.morphologyEx(bw, cv2.MORPH_OPEN, h_kernel, iterations=2)

h_contours, _ = cv2.findContours(
    h_lines, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE
)

h_positions = []
for cnt in h_contours:
    x, y, w_line, h_line = cv2.boundingRect(cnt)
    if w_line > tw * 0.6:
        h_positions.append(y)

# Add top & bottom
h_positions.extend([0, th - 1])
h_positions = sorted(set(h_positions))

# =====================================================
# STEP 5: DETECT VERTICAL LINES
# =====================================================
v_kernel = cv2.getStructuringElement(
    cv2.MORPH_RECT, (1, int(th * 0.4))
)
v_lines = cv2.morphologyEx(bw, cv2.MORPH_OPEN, v_kernel, iterations=2)

v_contours, _ = cv2.findContours(
    v_lines, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE
)

v_positions = []
for cnt in v_contours:
    x, y, w_line, h_line = cv2.boundingRect(cnt)
    if h_line > th * 0.5:
        v_positions.append(x)

# Add left & right
v_positions.extend([0, tw - 1])
v_positions = sorted(set(v_positions))

# =====================================================
# STEP 6: EXTRACT CELLS + OCR
# =====================================================
table_data = []

for i in range(len(h_positions) - 1):
    row = []
    for j in range(len(v_positions) - 1):
        y1, y2 = h_positions[i], h_positions[i + 1]
        x1, x2 = v_positions[j], v_positions[j + 1]

        # Skip very small boxes
        if (y2 - y1) < 15 or (x2 - x1) < 15:
            row.append("")
            continue

        cell_img = table_img[y1:y2, x1:x2]

        ocr = reader.readtext(cell_img, detail=0, paragraph=True)
        text = " ".join(ocr).strip()

        row.append(text)
    table_data.append(row)

# =====================================================
# STEP 7: SAVE TO EXCEL
# =====================================================
df = pd.DataFrame(table_data)

wb = Workbook()
ws = wb.active
ws.title = "Invoice Table"

for r in df.itertuples(index=False):
    ws.append(list(r))

wb.save(OUTPUT_EXCEL)

print("SUCCESS → Excel saved as:", OUTPUT_EXCEL)
