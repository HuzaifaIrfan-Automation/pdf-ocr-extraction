from pathlib import Path


import pytesseract

from PIL import Image
import os
import cv2
import pytesseract
from PIL import Image


import fitz  # PyMuPDF


# convert PDF to images

input_pdf_file_path="example.pdf"


def convert_pdf_to_images(pdf_path):
    file_name = Path(pdf_path).name
    doc = fitz.open(pdf_path)

    output_folder = Path(os.getcwd()) / "output" / file_name
    output_folder.mkdir(parents=True, exist_ok=True)

    for page_num in range(len(doc)):
        page = doc[page_num]
        pix = page.get_pixmap(dpi=500)  # render at 500 dpi
        pix.save(f"{output_folder}/{page_num+1}.png")

    return output_folder

output_folder = convert_pdf_to_images(input_pdf_file_path)
print(f"Converted PDF pages saved to: {output_folder}")


output_file_page_1 = f"{output_folder}/1.png"




data_map = {
    "dipendente_nome": {"type": str, "bbox": (544, 622, 1886, 696)},
    "codice_fiscale": {"type": str, "bbox": (1984, 340, 2671, 402)},
    "matricola_inps": {"type": str, "bbox": (2874, 208, 3244, 270)},
    "qualifica": {"type": str, "bbox": (700, 750, 1300, 850)},
    "mansione": {"type": str, "bbox": (1310, 750, 1882, 850)},
    "livello": {"type": int, "bbox": (1916, 750, 2190, 850)},
    "data_assunzione": {"type": str, "bbox": (126, 920, 392, 990)},
    "data_cessazione": {"type": str, "bbox": (410, 920, 686, 990)},
    "mese_retribuito": {"type": str, "bbox": (2020, 470, 2660, 580)},
    "anno": {"type": int, "bbox": (2910, 490, 3120, 580)},
    "totale_competenze": {"type": str, "bbox": (144, 4378, 502, 4440)},
    "totale_trattenute": {"type": str, "bbox": (2428, 5346, 2747, 5420)},
    "netto_in_busta": {"type": str, "bbox": (3290, 5338, 3880, 5410)},
    "imponibile_fiscale": {"type": str, "bbox": (146, 4512, 492, 4580)},
    "ritenute_inps": {"type": str, "bbox": (900, 4370, 1274, 4438)},
    "tfr_mese": {"type": str, "bbox": (1664, 4228, 2026, 4298)},
}


# Crop the image based on the bounding box

def crop_image(name, image_path, bbox):
    file_dir = Path(image_path).parent

    output_folder = Path(image_path).parent / "cropped"
    output_folder.mkdir(parents=True, exist_ok=True)
    output_file_path = f"{output_folder}/{name}.png"

    with Image.open(image_path) as img:
        cropped = img.crop(bbox)
        cropped.save(output_file_path)

    return output_file_path


cropped_image_paths = {}

for key, value in data_map.items():
    cropped_image_path = crop_image(key, output_file_page_1, value["bbox"])
    cropped_image_paths[key] = cropped_image_path


# Read and process cropped images with OCR

for key, cropped_image_path in cropped_image_paths.items():

    # Load image
    img = cv2.imread(cropped_image_path)

    # Convert to grayscale
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

    # Threshold (binarize)
    _, thresh = cv2.threshold(gray, 150, 255, cv2.THRESH_BINARY)

    # Optional: denoise
    thresh = cv2.medianBlur(thresh, 3)

    # OCR with multiple languages (example: English + Arabic + Chinese + German + Japanese + Russian + Urdu)
    # text = pytesseract.image_to_string(img_rgb, lang="eng+ara+chi_sim+deu+jpn+rus+urd").strip()
    text = pytesseract.image_to_string(thresh, lang="eng").strip()


    if data_map[key]["type"] == int:
        # Restrict to digits only
        custom_config = r'--oem 3 --psm 6 -c tessedit_char_whitelist=0123456789'

        text = pytesseract.image_to_string(thresh, config=custom_config).strip()


    # Print the extracted text

    print(f"---- {key}")
    print(text)

