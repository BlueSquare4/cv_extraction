from flask import Flask, render_template, request, send_file, url_for, redirect, session
import os
import pdfplumber
from docx import Document
import pandas as pd
import re
import tempfile

app = Flask(__name__)
app.secret_key = 'your_secret_key_here'  # Replace with your actual secret key

# Define utility functions for text extraction
def extract_text_from_pdf(pdf_path):
    text = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text() or ""
            text.append(page_text.replace('\n', ' '))  # Replace newline characters with spaces
    return " ".join(text)  # Join all text parts with a space to ensure words do not get concatenated directly


def extract_text_from_docx(docx_path):
    doc = Document(docx_path)
    text = " ".join(paragraph.text for paragraph in doc.paragraphs if paragraph.text)  # Use a space to join paragraphs
    return text

def extract_emails_phones_and_skill(text):
    email_pattern = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,7}\b'
    phone_pattern = r'\b\d{10}\b'  # Adjust as needed for expected formats
    skills_list = [
    'html', 'css', 'javascript', 'react', 'angular', 'vue.js', 'typescript', 'sass', 'less', 
    'bootstrap', 'foundation', 'ajax', 'jquery', 'webpack', 'gulp', 'babel', 'responsive design', 
    'accessibility', 'java', 'python', 'node.js', 'ruby', 'ruby on rails', 'php', 'c#', '.net', 
    'asp.net', 'spring boot', 'django', 'flask', 'sql', 'nosql', 'mongodb', 'mysql', 'postgresql', 
    'oracle', 'rest apis', 'graphql', 'api design', 'microservices', 'docker', 'kubernetes', 
    'matlab', 'excel', 'advanced excel', 'pivot tables', 'vlookup', 'macros', 'tableau', 'power bi', 
    'sas', 'spss', 'data analysis', 'statistical analysis', 'machine learning', 'deep learning', 
    'ai concepts', 'data visualization', 'big data technologies', 'apache hadoop', 'apache spark', 
    'data warehousing', 'etl processes', 'predictive analytics', 'financial modeling', 
    'business analysis', 'market analysis'
]

    emails = re.findall(email_pattern, text)
    phones = re.findall(phone_pattern, text)
    skills_found = [skill for skill in skills_list if skill.lower() in text.lower()]

    return emails, phones, skills_found

# Route for file uploads and processing
@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        files = request.files.getlist('files')  # Get a list of files
        if not files:
            return "No files selected"

        results = []
        for file in files:
            if file and (file.filename.endswith('.pdf') or file.filename.endswith('.docx')):
                temp_dir = tempfile.mkdtemp()
                filepath = os.path.join(temp_dir, file.filename)
                file.save(filepath)

                if file.filename.endswith('.pdf'):
                    text = extract_text_from_pdf(filepath)
                else:
                    text = extract_text_from_docx(filepath)

                emails, phones, skills = extract_emails_phones_and_skill(text)
                results.append({
                    "Filename": file.filename,
                    "Emails": ", ".join(set(emails)),
                    "Phones": ", ".join(set(phones)),
                    "Skills": ", ".join(skills),
                    "Text": text
                })

        df = pd.DataFrame(results)
        output_filename = os.path.join(temp_dir, 'extracted_info.xlsx')
        df.to_excel(output_filename, index=False)
        session['download_path'] = output_filename

        return render_template('index.html', table=df.to_html(classes='data', header="true"), filename='extracted_info.xlsx')
    
    return render_template('index.html', table=None)

# Route to download the Excel file
@app.route('/download')
def download_file():
    path = session.get('download_path', None)
    if path:
        return send_file(path, as_attachment=True)
    return "No file available to download."

if __name__ == '__main__':
    app.run(debug=True)
