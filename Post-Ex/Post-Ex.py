import PyPDF2
import re
import os
import openpyxl

current_directory = os.getcwd()

# List all files in the current directory
files = os.listdir(current_directory)
pdf_files = [file for file in files if file.lower().endswith('.pdf')]

wb = openpyxl.Workbook()
ws = wb.active
ws.append(['Name', 'Address', 'Contact', 'Destination','Tracking_no','Weight','COD_amount'])

def extract_info_from_text(text):
    # Initialize variables to store extracted information
    name = ""
    address = ""
    contact = ""
    tracking_no = ""
    destination = ""
    weight = ""
    cod_amount = ""
    
    # Extracting Name
    name_match = re.search(r"Receiver\s*:\s*(.*?)\s", text)
    if name_match:
        name = name_match.group(1).strip()

    # Extracting Full Address
    address_match = re.search(r"Address\s*:\s*(.*?)(?=Ye\s*parcel)", text, re.DOTALL)
    if address_match:
        address = address_match.group(1).strip()

    # Extracting Contact #
    contact_match = re.search(r"Phone\s*#\s*1\s*:\s*(.*?)\s", text)
    if contact_match:
        contact = contact_match.group(1).strip()

    # Extracting Tracking No
    tracking_no_match = re.search(r"Tracking\s*ID\s*:\s*(.*?)\s", text)
    if tracking_no_match:
        tracking_no = tracking_no_match.group(1).strip()

    # Extracting Destination
    destination_match = re.search(r"Destination\s*:\s*(.*?)\s", text)
    if destination_match:
        destination = destination_match.group(1).strip()

    # Extracting Weight
    weight_match = re.search(r"Weight\s*:\s*([\d.,]+)\s", text)
    if weight_match:
        weight = weight_match.group(1).strip()

    # Extracting COD Amount
    cod_amount_match = re.search(r"Rs\.\s*(\d[,\d]*)", text)
    if cod_amount_match:
        cod_amount = cod_amount_match.group(1).strip()

    return name, address, contact, tracking_no, destination, weight, cod_amount
# Iterate over the PDF files
for pdf_file in pdf_files:
    a=PyPDF2.PdfReader(pdf_file)
    NumPages = len(a.pages)
    found_courier_copy = False
        # Iterate over all pages of the PDF
    for page_num in range(NumPages):
        page = a.pages[page_num]
        text = page.extract_text() 
        name, address, contact, tracking_no, destination, weight, cod_amount=extract_info_from_text(text)
        ws.append([name, address, contact, destination,tracking_no,weight,cod_amount])
        print(address)
            

wb.save(os.path.join(current_directory, 'testt.xlsx'))