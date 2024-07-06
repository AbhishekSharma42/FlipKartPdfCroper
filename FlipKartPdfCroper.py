import os
from tkinter import Tk, Label, Button, Listbox, filedialog, messagebox
from PyPDF2 import PdfReader, PdfWriter
from PIL import Image
import pytesseract
from pdf2image import convert_from_path
from datetime import datetime

pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

def merge_pdfs(pdf_list, output):
    pdf_writer = PdfWriter()
    for pdf in pdf_list:
        pdf_reader = PdfReader(pdf)
        for page_num in range(len(pdf_reader.pages)):
            page = pdf_reader.pages[page_num]
            pdf_writer.add_page(page)
    with open(output, 'wb') as out:
        pdf_writer.write(out)

def crop_pdf(input_pdf, output_pdf, crop_area):
    pdf_reader = PdfReader(input_pdf)
    pdf_writer = PdfWriter()
    for page_num in range(len(pdf_reader.pages)):
        page = pdf_reader.pages[page_num]
        page.mediabox.lower_left = (crop_area['lower_left_x'], crop_area['lower_left_y'])
        page.mediabox.upper_right = (crop_area['upper_right_x'], crop_area['upper_right_y'])
        pdf_writer.add_page(page)
    with open(output_pdf, 'wb') as out:
        pdf_writer.write(out)

def extract_text_from_pdf(input_pdf):
    pages = convert_from_path(input_pdf,poppler_path=r'C:\pdfCroper\poppler-24.02.0\Library\bin')
    text = ""
    for page in pages:
        text += pytesseract.image_to_string(page)
    return text

def sort_pdfs_by_sku(pdf_list):
    sku_pdf_map = {}
    for pdf in pdf_list:
        text = extract_text_from_pdf(pdf)
        sku = extract_sku(text)
        sku_pdf_map[sku] = pdf
    sorted_pdfs = [sku_pdf_map[sku] for sku in sorted(sku_pdf_map)]
    return sorted_pdfs

def extract_sku(text):
    # Implement your logic to extract SKU from text
    # This is a placeholder implementation
    lines = text.split('\n')
    for line in lines:
        if "SKU" in line:
            return line.split()[-1]
    return "Unknown"

def select_files():
    files = filedialog.askopenfilenames(filetypes=[("PDF files", "*.pdf")])
    if files:
        pdf_listbox.delete(0, 'end')
        for file in files:
            pdf_listbox.insert('end', file)
        global pdf_files
        pdf_files = files

def process_files():
    if not pdf_files:
        messagebox.showerror("Error", "No PDF files selected")
        return
    
    current_datetime = datetime.now().strftime("%Y%m%d%H%M%S")
    merged_pdf = f'merged_{current_datetime}.pdf'
    cropped_pdf = f'cropped_{current_datetime}.pdf'
    sorted_merged_pdf = f'sorted_merged_{current_datetime}.pdf'
    
    merge_pdfs(pdf_files, merged_pdf)
    
    # Define crop area
    crop_area = {
        'lower_left_x': 185,
        'lower_left_y': 457,
        'upper_right_x': 410,
        'upper_right_y': 820
    }
    
    crop_pdf(merged_pdf, cropped_pdf, crop_area)
    
    sorted_pdfs = sort_pdfs_by_sku(pdf_files)
    merge_pdfs(sorted_pdfs, sorted_merged_pdf)
    
    messagebox.showinfo("Success", f"Your PDF Files Processed Successfully Cropped : {cropped_pdf}\n")

# Create the main window
root = Tk()
root.title("PDF Processor")

# Create and place widgets
Label(root, text="Select PDF files:").grid(row=0, column=0, padx=10, pady=10)

pdf_listbox = Listbox(root, width=50, height=10)
pdf_listbox.grid(row=1, column=0, padx=10, pady=10)

select_button = Button(root, text="Select PDFs", command=select_files)
select_button.grid(row=2, column=0, padx=10, pady=10)

process_button = Button(root, text="Process PDFs", command=process_files)
process_button.grid(row=3, column=0, padx=10, pady=10)

# Initialize the list of selected PDF files
pdf_files = []

# Run the application
root.mainloop()