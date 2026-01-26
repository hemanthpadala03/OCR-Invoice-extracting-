import cv2
import matplotlib.pyplot as plt


IMAGE_PATH = r"C:\Drive_d\Python\F-AI\T4\Input\images\ForwardInvoice_ORD46618691136_page_2.png"


img = cv2.imread(IMAGE_PATH)
img = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)
h, w, _ = img.shape

print(f"Image Width: {w}px, Height: {h}px")
print("Click on column boundaries (left → right). Close window when done.\n")


plt.figure(figsize=(8, 12))
plt.imshow(img)
plt.title("Click column boundaries (X only). Close window when finished.")
plt.axis("on")

points = plt.ginput(n=-1, timeout=0)
plt.close()

normalized_x = []

for (x, y) in points:
    norm = round(x / w, 4)
    normalized_x.append(norm)


normalized_x = sorted(normalized_x)


print("Captured Column X Coordinates (normalized by width):\n")

for i, x in enumerate(normalized_x, 1):
    print(f"Column {i}: {x}W")

print("\nRaw list (for direct reuse):")
print(normalized_x)
