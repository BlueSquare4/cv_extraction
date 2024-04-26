import os
import re
import pandas as pd
import pdfplumber
from docx import Document

def extract_text_from_pdf(pdf_path):
    text = ""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text += page.extract_text() or ""
    return text

def extract_text_from_docx(docx_path):
    doc = Document(docx_path)
    text = "\n".join([paragraph.text for paragraph in doc.paragraphs])
    return text

def extract_emails_and_phones(text):
    email_pattern = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,7}\b'
    phone_pattern = r'\b\d{10}\b'  # Adjust as needed for expected formats
    emails = re.findall(email_pattern, text)
    phones = re.findall(phone_pattern, text)
    return emails, phones

def process_files(directory):
    data = []
    for filename in os.listdir(directory):
        if filename.endswith(".pdf") or filename.endswith(".docx"):
            file_path = os.path.join(directory, filename)
            if filename.endswith(".pdf"):
                text = extract_text_from_pdf(file_path)
            else:
                text = extract_text_from_docx(file_path)

            emails, phones = extract_emails_and_phones(text)
            data.append({
                "Filename": filename,
                "Emails": ", ".join(set(emails)),
                "Phones": ", ".join(set(phones)),
                "Text": text
            })

    return data

# Specify the directory containing the CVs
directory = "Sample2"

# Process the files and get the data
extracted_data = process_files(directory)

# Create a DataFrame and write to an Excel file
df = pd.DataFrame(extracted_data)
df.to_excel("extracted_info.xlsx", index=False)

