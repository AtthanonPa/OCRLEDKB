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
from tkinter import filedialog, messagebox, ttk
import sys
import re
from shutil import copyfile
import subprocess
import platform

# Debugging mode toggle
DEBUG_MODE = False

# Set the path to the Tesseract executable (update as per your system configuration)
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# Fallback crop configuration percentages
BOX_QTY_FALLBACK_TOP = 0.87
BOX_QTY_FALLBACK_BOTTOM = 0.92
BOX_QTY_FALLBACK_LEFT = 0.73
BOX_QTY_FALLBACK_RIGHT = 0.87

def install_poppler():
    """
    Automatically install Poppler if not found.
    """
    if platform.system() == "Windows":
        poppler_url = "https://poppler.freedesktop.org/poppler-24.12.0.tar.xz"
        messagebox.showinfo("Poppler Required", "Poppler is required to process PDFs. Downloading and installing...")
        try:
            import requests, zipfile, io
            r = requests.get(poppler_url)
            z = zipfile.ZipFile(io.BytesIO(r.content))
            z.extractall("poppler")
            os.environ["PATH"] += os.pathsep + os.path.abspath("poppler/poppler-22.12.0-0/bin")
            messagebox.showinfo("Success", "Poppler installed successfully.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to install Poppler: {e}")
    else:
        messagebox.showerror("Unsupported OS", "Automatic Poppler installation is only supported on Windows.")

try:
    convert_from_path  # Check if Poppler is available
except Exception:
    install_poppler()

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

def extract_qr_and_box_qty_from_pdf(pdf_path, progress_callback=None):
    try:
        pages = convert_from_path(pdf_path, dpi=300)
        all_qr_data = []
        total_pages = len(pages)
        for idx, page in enumerate(pages):
            if progress_callback:
                progress_callback(idx + 1, total_pages)
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

def update_progress_bar(progress, total):
    percent = int((progress / total) * 100)
    progress_var.set(percent)
    progress_label.config(text=f"Progress: {percent}%")
    progress_bar.update()

def run_extraction():
    file_path = input_entry.get()
    output_path = output_entry.get()
    if not file_path or not output_path:
        messagebox.showerror("Error", "Please select input file and output location.")
        return

    if file_path.lower().endswith((".png", ".jpg", ".jpeg", ".tiff", ".bmp")):
        result = extract_qr_and_box_qty_from_image(file_path)
    elif file_path.lower().endswith(".pdf"):
        result = extract_qr_and_box_qty_from_pdf(file_path, update_progress_bar)
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

# Progress Bar
progress_var = tk.IntVar()
progress_bar = ttk.Progressbar(root, variable=progress_var, maximum=100)
progress_bar.grid(row=2, column=0, columnspan=3, padx=10, pady=10, sticky="we")
progress_label = tk.Label(root, text="Progress: 0%")
progress_label.grid(row=3, column=0, columnspan=3, padx=10, pady=5)
tk.Label(root, text="v 1.15 Develop by DX Engineer").grid(row=5, column=1, padx=0, pady=0)

# Run Button
tk.Button(root, text="Run", command=run_extraction, bg="green", fg="white").grid(row=4, column=1, padx=10, pady=20)

root.mainloop()
