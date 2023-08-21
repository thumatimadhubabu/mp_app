import os
import re
import tkinter as tk
from tkinter import filedialog
import pytesseract
from PIL import Image
import pandas as pd
import fitz
import cv2

pytesseract.pytesseract.tesseract_cmd = r'C:/Program Files/Tesseract-OCR/tesseract.exe'


def clean_text(text):
    # Remove unwanted symbols and characters using regular expressions
    cleaned_text = re.sub(r'[^A-Za-z0-9\s]', '', text)
    return cleaned_text

def ocr_to_excel(image_path, max_pages=None):
    _, file_extension = os.path.splitext(image_path)

    # Handle image files using PIL
    if file_extension.lower() in ['.jpeg', '.jpg', '.png', '.gif', '.bmp']:
        # Perform OCR on the image using Tesseract
        extracted_text = pytesseract.image_to_string(Image.open(image_path), lang='eng')

    # Handle PDF files using PyMuPDF
    elif file_extension.lower() == '.pdf':
        pdf_doc = fitz.open(image_path)
        extracted_text = ""
        for page_num in range(min(pdf_doc.page_count, max_pages) if max_pages else pdf_doc.page_count):
            page = pdf_doc[page_num]

            # Convert the PDF page to an image (PNG) using PyMuPDF
            pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))  # Increase the resolution for better OCR results

            # Create a PIL image from the pixmap
            image = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

            temp_image_path = f"temp_page_{page_num}.png"
            image.save(temp_image_path, format="PNG")  # Save the image as PNG

            # Perform OCR on the extracted image using Tesseract
            page_text = pytesseract.image_to_string(Image.open(temp_image_path), lang='eng')
            extracted_text += page_text + '\n'

            os.remove(temp_image_path)  # Remove the temporary image

        pdf_doc.close()
    else:
        print(f"Unsupported file format: {file_extension}")
        return

    # Clean the extracted text
    cleaned_text = clean_text(extracted_text)

    # Split the cleaned text by lines to separate rows
    rows = cleaned_text.split('\n')

    # Create a list to store the extracted data as dictionaries
    data_list = []
    for row in rows:
        # Split each row by spaces to separate columns
        columns = row.split()
        data_dict = {}
        for i, col in enumerate(columns):
            # Use column index as the key for the dictionary
            data_dict[f"Column{i + 1}"] = col
        data_list.append(data_dict)

    # Convert the list of dictionaries to a DataFrame
    df = pd.DataFrame(data_list)

    # Save the DataFrame to an Excel file
    excel_file_path = image_path.replace(file_extension, '.xlsx')
    df.to_excel(excel_file_path, index=False)
    return excel_file_path

def select_files():
    file_paths = filedialog.askopenfilenames(
        filetypes=[('Image and PDF Files', '*.jpeg;*.jpg;*.png;*.gif;*.bmp;*.pdf')])
    if file_paths:
        max_pages = int(input("Enter the number of pages to OCR (Enter 0 or leave blank to process all pages): "))
        for file_path in file_paths:
            output_path = ocr_to_excel(file_path, max_pages)
            print(f"Output saved in: {output_path}")

if __name__ == "__main__":
    root = tk.Tk()
    root.title("OCR to Excel")

    select_button = tk.Button(root, text="Select Files", command=select_files)
    select_button.pack(pady=10)

    root.mainloop()
