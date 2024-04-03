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
ws.append(['Name', 'Address', 'Contact', 'Destination','Tracking_no','Weight','COD_amount','Booking_Date'])

# Iterate over the PDF files
for pdf_file in pdf_files:
    a=PyPDF2.PdfReader(pdf_file)
    NumPages = len(a.pages)
    found_courier_copy = False
        # Iterate over all pages of the PDF
    for page_num in range(NumPages):
        page = a.pages[page_num]
        text = page.extract_text()
        chunks = re.split(r'\bUAN\b', text)
        #print(text.lower())
        if "courier copy" in text.lower():
            found_courier_copy = True
            break  # Break the loop if "courier copy" is found
        for chunk in chunks:
            
            
    # Initialize variables to store information
            name = ""
            address = ""
            contact = ""        
            lines = chunk.split('\n')
            first_address_found = False
            capture_address = False
    # Iterate through each line to find the required information
            for line in lines:
                if "Name :" in line:
                    name = line.split(":")[1].strip()
                elif "Address :" in line and not first_address_found:
                    address = line.split(":")[1].strip()
                    first_address_found = True
                    capture_address = True
                elif "Contact #:" in line:
                    capture_address = False
                    contact = line.split(":")[1].strip()
                elif capture_address:
            # Concatenate lines to capture multi-line address
                    address += " " + line.strip()

            consignee_info_regex = r"Consignee Information\s+Name\s*:\s*([^\n]+)\s+Address\s*:\s*([^\n]+)\s+Contact #\s*:\s*([\d,\s]+)"
            destination_regex = r"Destination\s*:\s*([^P]+)\s*Pieces"
            tracking_info_regex = r"Tracking No\s*:\s*([\d]+)"
            weight_regex = r"Weight\s*:\s*([\d.,]+)\s*\("
            cod_amount_regex = r"COD\s*Amount\s*:\s*PKR\s*([\d.,]+)"
            date_regex = r"Booking\s*Date\s*:\s*([\d-]+)"

            tracking_info_match = re.search(tracking_info_regex, chunk)
            weight_match = re.search(weight_regex, chunk)
            cod_amount_match = re.search(cod_amount_regex, chunk)
            date_match = re.search(date_regex, chunk)
            destination_match = re.search(destination_regex, chunk)
            print("---------------")
            print("Name:", name)
            print("Address:", address)
            print("Contact #:", contact)
                   
            if destination_match:
                destination = destination_match.group(1).strip()
            if tracking_info_match:
                tracking_no = tracking_info_match.group(1)
            if weight_match:
                weight = weight_match.group(1)
            if cod_amount_match:
                cod_amount = cod_amount_match.group(1)
            if date_match:
                booking_date = date_match.group(1)
            if name=="":
                continue
            else:
                ws.append([name, address, contact, destination,tracking_no,weight,cod_amount,booking_date])

wb.save(os.path.join(current_directory, 'testt.xlsx'))


