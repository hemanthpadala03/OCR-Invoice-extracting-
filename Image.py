import os
from pdf2image import convert_from_path

INPUT_DIR = r"C:\Drive_d\Python\F-AI\T4\Input"
OUTPUT_DIR = os.path.join(INPUT_DIR, "images")

os.makedirs(OUTPUT_DIR, exist_ok=True)

for file in os.listdir(INPUT_DIR):
    if not file.lower().endswith(".pdf"):
        continue

    pdf_path = os.path.join(INPUT_DIR, file)
    name = os.path.splitext(file)[0]

    pages = convert_from_path(pdf_path, dpi=200)

    for i, page in enumerate(pages, start=1):
        page.save(
            os.path.join(OUTPUT_DIR, f"{name}_page_{i}.png"),
            "PNG"
        )

    print(f"Converted: {file}")
