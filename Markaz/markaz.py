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
    name = ""
    phone = ""
    address = ""
    tracking_no = ""
    weight = ""
    cod_amount = ""
    date = ""
    destination = ""

    # Extracting Date
    date_match = re.search(r"Shipping\s+Date:\s*(\d{2}/\d{2}/\d{4})", text)
    if date_match:
        date = date_match.group(1)

    # Extracting Name
    name_match = re.search(r"Name:\s*([^\n]+)", text)
    if name_match:
        name = name_match.group(1).strip()

    # Extracting Phone
    phone_match = re.search(r"Phone:\s*([^\n]+)", text)
    if phone_match:
        phone = phone_match.group(1).strip()

    # Extracting Address
    address_match = re.search(r"Address:\s*([\s\S]+?)(?=Mashoor\s*jagah|Description:)", text)
    if address_match:
        address = address_match.group(1).strip()

    # Extracting Destination
    destination_match = re.search(r"Destination\s+City:\s*([^\n]+)", text)
    if destination_match:
        destination = destination_match.group(1).strip()

    # Extracting Tracking No
    tracking_no_match = re.search(r"([A-Za-z0-9]+)\s*\n\s*Description:", text)
    if tracking_no_match:
        tracking_no = tracking_no_match.group(1)

    # Extracting Weight
    weight_match = re.search(r"Weight:\s*([\d.,]+)", text)
    if weight_match:
        weight = weight_match.group(1).strip()

    # Extracting COD Amount
    cod_amount_match = re.search(r"Amount:\s*(\d+)", text)
    if cod_amount_match:
        cod_amount = cod_amount_match.group(1).strip()

    return date, name, phone, address, destination, tracking_no, weight, cod_amount
# Iterate over the PDF files
for pdf_file in pdf_files:
    a=PyPDF2.PdfReader(pdf_file)
    NumPages = len(a.pages)
    found_courier_copy = False
        # Iterate over all pages of the PDF
    count=0
    for page_num in range(NumPages):
        page = a.pages[page_num]
        text = page.extract_text() 
        if "markaz supplier order sheet" in text.lower():
            continue
        else:
            date,name,contact , address,destination, tracking_no, weight, cod_amount=extract_info_from_text(text)
        ws.append([name, address, contact, destination,tracking_no,weight,cod_amount,date])
wb.save(os.path.join(current_directory, 'testt.xlsx'))
