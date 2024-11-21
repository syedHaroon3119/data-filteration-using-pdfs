import os
import fitz  # PyMuPDF
import cv2
import numpy as np
from pyzbar.pyzbar import decode
import pandas as pd
from multiprocessing import Pool

# Function to extract QR from PDF
def extract_qr_from_pdf(pdf_path):
    doc = fitz.open(pdf_path)
    for page_number in range(len(doc)):
        page = doc.load_page(page_number)
        # Extract image from the page
        img_list = page.get_images(full=True)
        for img_index, img in enumerate(img_list):
            xref = img[0]
            base_image = doc.extract_image(xref)
            image_bytes = base_image["image"]
            
            # Convert the extracted image bytes into an OpenCV image
            image = cv2.imdecode(np.frombuffer(image_bytes, np.uint8), cv2.IMREAD_COLOR)
            
            # Decode QR codes in the image
            qr_codes = decode(image)
            for qr in qr_codes:
                qr_data = qr.data.decode("utf-8")  # Get QR data as string
                return qr_data  # Return the first QR code found
    return None  # If no QR code found

# Function to process a single PDF file
def process_pdf(pdf):
    qr_data = extract_qr_from_pdf(pdf)
    if qr_data:
        # Parse the QR data (which is a JSON string) into a dictionary
        import json
        data = json.loads(qr_data)
        
        # Extract necessary information from the QR data
        extracted_data = {
            "ABHA Number": data.get("hidn"),
            "HID": data.get("hid"),
            "Name": data.get("name"),
            "Gender": data.get("gender"),
            "DOB": data.get("dob"),
            "District": data.get("district_name"),
            "State": data.get("state name"),
            "Mobile": data.get("mobile"),
            "Address": data.get("address"),
        }
        return extracted_data
    return None

# Main function to handle folder processing and multiprocessing
if __name__ == "__main__":
    # Folder containing your PDFs (relative path to the script)
    pdf_folder = "pdfs1"  # This should be the folder name where PDFs are stored

    # Get all PDF files in the folder
    pdf_files = [os.path.join(pdf_folder, f) for f in os.listdir(pdf_folder) if f.lower().endswith(".pdf")]

    # Use multiprocessing to process PDFs in parallel
    with Pool(processes=4) as pool:  # Adjust 'processes' as per your CPU cores
        results = pool.map(process_pdf, pdf_files)

    # Filter out None results (PDFs with no QR code) and create DataFrame
    filtered_results = [r for r in results if r]
    df = pd.DataFrame(filtered_results)

    # Export the data to an Excel file
    df.to_excel("qr_data_sorted.xlsx", index=False)

    print("QR code data has been extracted and saved to 'qr_data_sorted.xlsx'.")
