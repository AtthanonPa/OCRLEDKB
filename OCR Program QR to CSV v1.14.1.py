# Open Source OCR Program with GUI
# This script uses Tesseract OCR, an open-source optical character recognition engine,
# to extract QR code data and numeric text below 'Box QTY' from images and PDF files,
# formats it, and exports it to a CSV file. A graphical user interface (GUI) allows for
# file selection and output location. Ensure Tesseract, pdf2image, and pyzbar libraries are installed.

import pytesseract
from PIL import Image
from pdf2image import convert_from_path
import cv2
import numpy as np
from pyzbar.pyzbar import decode
import os
import csv
import tkinter as tk
from tkinter import filedialog, messagebox
import sys
import re
from shutil import copyfile

# Debugging mode toggle
DEBUG_MODE = False

# Set the path to the Tesseract executable (update as per your system configuration)
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# Fallback crop configuration percentages
BOX_QTY_FALLBACK_TOP = 0.87
BOX_QTY_FALLBACK_BOTTOM = 0.92
BOX_QTY_FALLBACK_LEFT = 0.73
BOX_QTY_FALLBACK_RIGHT = 0.87

def preprocess_image(image):
    """
    Preprocess the image to improve OCR and QR code detection.
    Converts to grayscale and applies thresholding.
    """
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    _, thresh = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY | cv2.THRESH_OTSU)
    return thresh

def dynamic_crop_box_qty_area(image):
    """
    Dynamically finds 'BOX QTY' and crops a refined area below it.
    """
    try:
        ocr_data = pytesseract.image_to_data(image, config='--psm 6', output_type=pytesseract.Output.DICT)
        for i, text in enumerate(ocr_data['text']):
            if text and "BOX QTY" in text.upper():
                x, y, w, h = ocr_data['left'][i], ocr_data['top'][i], ocr_data['width'][i], ocr_data['height'][i]
                cropped_image = image[y + h + 5:y + h + 70, x:x + w + 20]
                return cropped_image
    except Exception as e:
        pass

    height, width = image.shape[:2]
    cropped_static = image[
        int(BOX_QTY_FALLBACK_TOP * height):int(BOX_QTY_FALLBACK_BOTTOM * height),
        int(BOX_QTY_FALLBACK_LEFT * width):int(BOX_QTY_FALLBACK_RIGHT * width)
    ]
    return cropped_static

def preprocess_for_box_qty(cropped_image):
    """
    Preprocess the cropped BOX QTY area for OCR using adaptive thresholding.
    """
    try:
        gray = cv2.cvtColor(cropped_image, cv2.COLOR_BGR2GRAY)
        resized = cv2.resize(gray, None, fx=3, fy=3, interpolation=cv2.INTER_LINEAR)
        blurred = cv2.GaussianBlur(resized, (5, 5), 0)
        _, thresh = cv2.threshold(blurred, 0, 255, cv2.THRESH_BINARY | cv2.THRESH_OTSU)
        return thresh
    except Exception as e:
        return cropped_image

def extract_box_qty_from_text(text):
    """
    Extracts the numeric value for 'Box QTY' using refined regex.
    """
    numbers = re.findall(r'\b\d+\b', text.strip())
    return numbers[0] if numbers else "N/A"

def extract_qr_and_box_qty(image):
    """
    Extracts QR code data and the numeric value for 'Box QTY' from the image.
    """
    preprocessed_image = preprocess_image(image)
    qr_data_list = [qr.data.decode('utf-8') for qr in decode(preprocessed_image)]

    cropped_box_qty_image = dynamic_crop_box_qty_area(image)
    box_qty = "N/A"
    if cropped_box_qty_image is not None:
        processed_box_qty_image = preprocess_for_box_qty(cropped_box_qty_image)
        box_qty_text = pytesseract.image_to_string(
            processed_box_qty_image, config='--psm 7 -c tessedit_char_whitelist=0123456789'
        )
        box_qty = extract_box_qty_from_text(box_qty_text)

    return qr_data_list, box_qty

def process_page(page):
    """
    Processes a single PDF page to extract QR codes and Box QTY.
    """
    image = cv2.cvtColor(np.array(page), cv2.COLOR_RGB2BGR)
    qr_data_list, box_qty = extract_qr_and_box_qty(image)
    result_data = []
    for qr_data in qr_data_list:
        qr_columns = qr_data.replace(":", "|").split('|')
        qr_columns.extend([box_qty, "P1"])
        result_data.append(qr_columns)
    return result_data

def extract_qr_and_box_qty_from_pdf(pdf_path):
    try:
        pages = convert_from_path(pdf_path, dpi=300)
        all_qr_data = []
        for page in pages:
            all_qr_data.extend(process_page(page))
        return all_qr_data
    except Exception as e:
        return [[f"Error: {e}"]]

def extract_qr_and_box_qty_from_image(image_path):
    try:
        image = cv2.imread(image_path)
        qr_data_list, box_qty = extract_qr_and_box_qty(image)
        result_data = []
        for qr_data in qr_data_list:
            qr_columns = qr_data.replace(":", "|").split('|')
            qr_columns.extend([box_qty, "P1"])
            result_data.append(qr_columns)
        return result_data
    except Exception as e:
        return [[f"Error: {e}"]]

def save_to_csv(output_path, data):
    try:
        with open(output_path, 'w', newline='', encoding='utf-8') as csv_file:
            writer = csv.writer(csv_file)
            writer.writerows(data)
        messagebox.showinfo("Success", f"Data saved to {output_path}")
    except Exception as e:
        messagebox.showerror("Error", f"Error saving file: {e}")

def browse_file():
    file_path = filedialog.askopenfilename(filetypes=[("PDF and Images", "*.pdf;*.png;*.jpg;*.jpeg;*.tiff;*.bmp")])
    if file_path:
        input_entry.delete(0, tk.END)
        input_entry.insert(0, file_path)

def browse_output():
    output_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV Files", "*.csv")])
    if output_path:
        output_entry.delete(0, tk.END)
        output_entry.insert(0, output_path)

def run_extraction():
    file_path = input_entry.get()
    output_path = output_entry.get()
    if not file_path or not output_path:
        messagebox.showerror("Error", "Please select input file and output location.")
        return

    if file_path.lower().endswith((".png", ".jpg", ".jpeg", ".tiff", ".bmp")):
        result = extract_qr_and_box_qty_from_image(file_path)
    elif file_path.lower().endswith(".pdf"):
        result = extract_qr_and_box_qty_from_pdf(file_path)
    else:
        messagebox.showerror("Error", "Unsupported file format.")
        return

    if result and any(len(row) > 0 for row in result):
        save_to_csv(output_path, result)
    else:
        messagebox.showwarning("No Data Found", "No QR code or BOX QTY data could be extracted.")

# GUI Setup
root = tk.Tk()
root.title("QR Code and Box QTY Reader")

# Input File Selection
tk.Label(root, text="Select Input File:").grid(row=0, column=0, padx=10, pady=10)
input_entry = tk.Entry(root, width=50)
input_entry.grid(row=0, column=1, padx=10, pady=10)
tk.Button(root, text="Browse", command=browse_file).grid(row=0, column=2, padx=10, pady=10)

# Output File Selection
tk.Label(root, text="Select Output File:").grid(row=1, column=0, padx=10, pady=10)
output_entry = tk.Entry(root, width=50)
output_entry.grid(row=1, column=1, padx=10, pady=10)
tk.Button(root, text="Browse", command=browse_output).grid(row=1, column=2, padx=10, pady=10)
tk.Label(root, text="v 1.14.1 Develop by DX Engineer").grid(row=3, column=1, padx=0, pady=0)

# Run Button
tk.Button(root, text="Run", command=run_extraction, bg="green", fg="white").grid(row=2, column=1, padx=10, pady=20)

root.mainloop()
