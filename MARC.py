import PyPDF2
import os
from re import search
from tkinter.filedialog import askdirectory

input("Press Enter to Select Input BEO folder.")
beo_folder = askdirectory()
input("Press Enter to Select Output MARC BEO folder.")
marc_beo_folder = askdirectory()

print("Hi, I'm M.A.R.C., Michael's Automated R2 Condenser.")
internet = input("Do you handle internet services at your location? Y/N\n")

while type(internet) == str:
    if internet.lower() == "y":
        print("Great, I will grab internet service information as well.\n")
        internet = True
    elif internet.lower() == "n":
        print("Great, I will only grab the audio visual information.\n")
        internet = False
    else:
        internet = input("That was not a vaild answer.\nDo you handle internet services at your location? Y/N\n")

print("Condensing BEOs the M.A.R.C.ey way...\n")

for beo in os.scandir(beo_folder):
    beo_name = os.path.basename(beo)
    beo = open(beo, 'rb')
    beo_pdf = PyPDF2.PdfFileReader(beo)
    pdf_writer = PyPDF2.PdfFileWriter()

    og_document_len = beo_pdf.getNumPages()
    tagged_number = []
    current_page = 0

    for page in range(1, og_document_len):
        page = beo_pdf.getPage(current_page)
        page_text = page.extractText()
        page_text = page_text.lower()
        page_tag = False
        
        if search("audio visual", page_text):
            page_tag = True
            
        if search("no av", page_text) or search("own av", page_text):
            page_tag = False

        if internet == True:
            if search("information technology", page_text):
                page_tag = True
            if search("internet service", page_text):
                page_tag = True
            
        if page_tag == True:
            beo_number = search('event order #:(\d+)', page_text[10:40])
            beo_number = beo_number.group(1)
            if beo_number not in tagged_number:
                tagged_number.append(beo_number)
        
        current_page += 1

    current_page = 0

    for page in range(1, og_document_len):
        page = beo_pdf.getPage(current_page)
        page_text = page.extractText()

        for number in tagged_number:
            if search(number, page_text):
                pdf_writer.addPage(page)
                break

        current_page += 1

    marc_name = "MARC BEO - " + beo_name
    marc_name_path = os.path.join(marc_beo_folder, marc_name)
    marc_beo = open(marc_name_path, 'wb')
    pdf_writer.write(marc_beo)
    print(f"File was saved to {marc_name_path}")
    beo.close()
    marc_beo.close()

input("\nPress Enter to exit.")
