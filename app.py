from flask import Flask, request, send_file, render_template
import os
import io
import zipfile
import csv
from certificate import *
from docx import Document
from docx2pdf import convert
from openpyxl import Workbook, load_workbook
import pythoncom

app = Flask(__name__)

# Ensure the necessary directories exist
os.makedirs('uploads', exist_ok=True)
os.makedirs('Output/Doc', exist_ok=True)
os.makedirs('Output/Pdf', exist_ok=True)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/generate_certificate', methods=['POST'])
def generate_certificate():
    ambassador = request.form['ambassador']
    event = request.form['event']
    csv_file = request.files['csv_file']
    filename = csv_file.filename

    if not csv_file or not filename.endswith('.csv'):
        return "Invalid file format. Please upload a CSV file."

    csv_path = os.path.join('uploads', filename)
    csv_file.save(csv_path)

    participants = get_participants(csv_path)

    # Initialize COM
    pythoncom.CoInitialize()
    
    zip_buffer = create_docx_files('Data/Event Certificate Template.docx', participants, event, ambassador)

    # Uninitialize COM
    pythoncom.CoUninitialize()

    return send_file(zip_buffer, as_attachment=True, download_name='certificates.zip', mimetype='application/zip')

def get_participants(file_path):
    data = []
    with open(file_path, mode="r", encoding='iso-8859-1') as file:
        csv_reader = csv.DictReader(file)
        for row in csv_reader:
            data.append(row)
    return data

def create_docx_files(template_path, participants, event, ambassador):
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, 'a', zipfile.ZIP_DEFLATED) as zip_file:
        for participant in participants:
            doc = Document(template_path)
            name = participant['Name']
            email = participant['Email']

            replace_participant_name(doc, name)
            replace_event_name(doc, event)
            replace_ambassador_name(doc, ambassador)

            doc_path = f'Output/Doc/{name}.docx'
            pdf_path = f'Output/Pdf/{name}.pdf'

            doc.save(doc_path)
            convert(doc_path, pdf_path)

            with open(pdf_path, 'rb') as pdf_file:
                zip_file.writestr(f'{name}.pdf', pdf_file.read())

    zip_buffer.seek(0)
    return zip_buffer

if __name__ == '__main__':
    app.run(debug=True)
