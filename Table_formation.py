import cv2
import easyocr
import numpy as np

# =====================================================
# CONFIG
# =====================================================
IMAGE_PATH = r"C:\Drive_d\Python\F-AI\T4\Input\images\ForwardInvoice_ORD46618691136_page_2.png"

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
# STEP 3: PREPROCESS
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

detect_h = cv2.morphologyEx(
    bw, cv2.MORPH_OPEN, h_kernel, iterations=2
)

h_contours, _ = cv2.findContours(
    detect_h, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE
)

# Draw horizontal lines (BLUE)
for cnt in h_contours:
    x, y, w_line, h_line = cv2.boundingRect(cnt)
    if w_line > tw * 0.6:
        cv2.line(
            table_img,
            (0, y),
            (tw, y),
            (255, 0, 0),
            2
        )

# =====================================================
# STEP 5: DETECT VERTICAL LINES (SAME SYSTEM)
# =====================================================
v_kernel = cv2.getStructuringElement(
    cv2.MORPH_RECT, (1, int(th * 0.4))
)

detect_v = cv2.morphologyEx(
    bw, cv2.MORPH_OPEN, v_kernel, iterations=2
)

v_contours, _ = cv2.findContours(
    detect_v, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE
)

# Draw vertical lines (RED)
for cnt in v_contours:
    x, y, w_line, h_line = cv2.boundingRect(cnt)
    if h_line > th * 0.5:
        cv2.line(
            table_img,
            (x, 0),
            (x, th),
            (0, 0, 255),
            2
        )

# =====================================================
# STEP 6: ADD BOTTOM BORDER (SAFETY)
# =====================================================
cv2.line(
    table_img,
    (0, th - 2),
    (tw, th - 2),
    (255, 0, 0),
    2
)

# =====================================================
# SHOW RESULT
# =====================================================
cv2.imshow("Table Grid (Pure Geometry)", table_img)
cv2.waitKey(0)
cv2.destroyAllWindows()
