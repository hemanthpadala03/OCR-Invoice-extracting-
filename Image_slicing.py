import cv2
import easyocr

IMAGE_PATH = r"C:\Drive_d\Python\F-AI\T4\Input\images\ForwardInvoice_ORD46618691136_page_2.png"

# Load image
img = cv2.imread(IMAGE_PATH)
h, w, _ = img.shape

reader = easyocr.Reader(['en'], gpu=False)

results = reader.readtext(img, detail=1)

state_y_max = None
for box, text, conf in results:
    if text.strip().lower() == "state":
        y_coords = [p[1] for p in box]
        state_y_max = int(max(y_coords))
        break

if state_y_max is None:
    raise RuntimeError("State not found")

lower_after_state = img[state_y_max + 10 : h, 0:w]

results_lower = reader.readtext(lower_after_state, detail=1)

amount_y_min = None
for box, text, conf in results_lower:
    if "amount in" in text.lower():
        y_coords = [p[1] for p in box]
        amount_y_min = int(min(y_coords))
        break

if amount_y_min is None:
    raise RuntimeError("Amount in not found")


middle_part = lower_after_state[0 : amount_y_min - 10, 0:w]


cv2.imshow("Middle Part (Table Region)", middle_part)
cv2.waitKey(0)
cv2.destroyAllWindows()
