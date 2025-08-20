from pathlib import Path


import pytesseract

from PIL import Image
import os


import fitz  # PyMuPDF


input_pdf_file_path="input/example.pdf"



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



bboxes = {
    "dipendente_nome": (544, 622, 1886, 696),
    "codice_fiscale": (1984, 340, 2671, 402),
    "matricola_inps": (0, 0, 10, 10),
    "qualifica_mansione": (0, 0, 10, 10),
    "livello":(0, 0, 10, 10),
    "data_assunzione": (0, 0, 10, 10),
    "data_cessazione":(0, 0, 10, 10),
    "mese_retribuito": (0, 0, 10, 10),
    "anno": (0, 0, 10, 10),
    "totale_competenze": (0, 0, 10, 10),
    "totale_trattenute":(0, 0, 10, 10),
    "netto_in_busta": (0, 0, 10, 10),
    "imponibile_fiscale":(0, 0, 10, 10),
    "ritenute_inps":(0, 0, 10, 10),
    "tfr_mese": (0, 0, 10, 10),
}

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

for name, bbox in bboxes.items():
    cropped_image_path = crop_image(name, output_file_page_1, bbox)
    cropped_image_paths[name] = cropped_image_path



for key, cropped_image_path in cropped_image_paths.items():
    # Open the cropped image
    cropped_image = Image.open(cropped_image_path)

    # Use pytesseract to do OCR on the cropped image
    text = pytesseract.image_to_string(cropped_image)

    # Print the extracted text

    print(f"---- {key}")
    print(text)

