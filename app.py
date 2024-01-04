from flask import Flask, render_template, request, send_from_directory
from pdf2docx import Converter
import os
from io import BytesIO
from docx import Document
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
import threading

app = Flask(__name__ , static_url_path='/static')

# Define the upload and converted file directories
UPLOAD_FOLDER = 'uploads'
CONVERTED_FOLDER = 'converted_files'

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['CONVERTED_FOLDER'] = CONVERTED_FOLDER

# Ensure the upload and converted directories exist
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(CONVERTED_FOLDER, exist_ok=True)

@app.route('/')
def start():
    return render_template('start.html')

@app.route('/convert_pdf_to_word')
def index():
    return render_template('index.html')

@app.route('/convert_pdf_to_word', methods=['POST'])
def convert():
    # Check if a file was uploaded
    if 'fileInput' not in request.files:
        return "No file provided."

    pdf_file = request.files['fileInput']

    # Check if the file has a name
    if pdf_file.filename == '':
        return "No selected file."

    # Save the uploaded PDF file with its original name
    pdf_file_path = os.path.join(app.config['UPLOAD_FOLDER'], pdf_file.filename)
    pdf_file.save(pdf_file_path)

    # Perform PDF to DOCX conversion
    docx_file_path = os.path.join(app.config['CONVERTED_FOLDER'], f'{os.path.splitext(pdf_file.filename)[0]}.docx')
    convert_pdf_to_docx(pdf_file_path, docx_file_path)

    # Provide the converted file for download
    return send_from_directory(app.config['CONVERTED_FOLDER'], f'{os.path.splitext(pdf_file.filename)[0]}.docx', as_attachment=True)


@app.route('/convert_word_to_pdf')
def word_to_pdf():
    return render_template('word_to_pdf.html')


@app.route('/convert_word_to_pdf', methods=['POST'])
def convert_to_pdf():
    # Check if a file was uploaded
    if 'fileInput' not in request.files:
        return "No file provided."

    pdf_file = request.files['fileInput']

    # Check if the file has a name
    if pdf_file.filename == '':
        return "No selected file."

    # Save the uploaded PDF file with its original name
    pdf_file_path = os.path.join(app.config['UPLOAD_FOLDER'], pdf_file.filename)
    pdf_file.save(pdf_file_path)

    # Perform DOCX to PDF conversion
    pdf_file_path_output = os.path.join(app.config['CONVERTED_FOLDER'], f'{os.path.splitext(pdf_file.filename)[0]}.pdf')
    convert_docx_to_pdf(pdf_file_path, pdf_file_path_output)

 

    # Provide the converted file for download
    return send_from_directory(app.config['CONVERTED_FOLDER'], f'{os.path.splitext(pdf_file.filename)[0]}.pdf', as_attachment=True)

def convert_docx_to_pdf(docx_path, pdf_path):
    if os.path.exists(pdf_path):
        try:
            os.remove(pdf_path)
        except Exception as e:
            return f"Error deleting existing file: {str(e)}"

    try:
        # Load the DOCX file
        doc = Document(docx_path)

        # Create a BytesIO object to save the PDF content
        pdf_output = BytesIO()

        # Create a PDF canvas
        pdf_canvas = canvas.Canvas(pdf_output, pagesize=letter)

        # Iterate through paragraphs in the DOCX file and add them to the PDF canvas
        for paragraph in doc.paragraphs:
            pdf_canvas.drawString(10, 800, paragraph.text)

        # Save the PDF content to the BytesIO object
        pdf_canvas.save()

        # Write the BytesIO content to the PDF file
        with open(pdf_path, 'wb') as pdf_file:
            pdf_file.write(pdf_output.getvalue())

        pdf_output.close()

    except Exception as e:
        return f"Error during conversion: {str(e)}"


def convert_pdf_to_docx(pdf_path, docx_path):
    # Check if the file already exists and delete it before conversion
    if os.path.exists(docx_path):
        try:
            os.remove(docx_path)
        except Exception as e:
            return f"Error deleting existing file: {str(e)}"

    try:
        cv = Converter(pdf_path)
        cv.convert(docx_path)
        cv.close()  # Close the Converter instance
    except Exception as e:
        return f"Error during conversion: {str(e)}"

if __name__ == '__main__':
    app.run(debug=True)