
import os
import fitz  # PyMuPDF
import shutil
from flask import Flask, request, send_file, render_template, redirect, url_for
from werkzeug.utils import secure_filename
from PyPDF2 import PdfReader, PdfWriter
from pdf2image import convert_from_path
from docx import Document
from PIL import Image
import uuid

app = Flask(__name__)
UPLOAD_FOLDER = 'static/uploads'
DOWNLOAD_FOLDER = 'static/downloads'
ALLOWED_EXTENSIONS = {'pdf'}

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(DOWNLOAD_FOLDER, exist_ok=True)

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def cleanup():
    for folder in [UPLOAD_FOLDER, DOWNLOAD_FOLDER]:
        for file in os.listdir(folder):
            os.remove(os.path.join(folder, file))

@app.route('/')
def home():
    return render_template('index.html')

@app.route('/process', methods=['POST'])
def process():
    cleanup()
    tool = request.form.get('tool')
    uploaded_files = request.files.getlist('file')
    options = request.form.to_dict()

    saved_files = []
    for f in uploaded_files:
        if f and allowed_file(f.filename):
            filename = secure_filename(f.filename)
            path = os.path.join(UPLOAD_FOLDER, filename)
            f.save(path)
            saved_files.append(path)

    output_file = os.path.join(DOWNLOAD_FOLDER, f"{uuid.uuid4().hex}.pdf")

    try:
        if tool == 'compress':
            doc = fitz.open(saved_files[0])
            doc.save(output_file, deflate=True, garbage=3)
            doc.close()
        elif tool == 'merge':
            merger = PdfWriter()
            for f in saved_files:
                reader = PdfReader(f)
                for page in reader.pages:
                    merger.add_page(page)
            with open(output_file, "wb") as f:
                merger.write(f)
        elif tool == 'split':
            reader = PdfReader(saved_files[0])
            for i, page in enumerate(reader.pages):
                writer = PdfWriter()
                writer.add_page(page)
                part_path = os.path.join(DOWNLOAD_FOLDER, f"{uuid.uuid4().hex}_page{i+1}.pdf")
                with open(part_path, "wb") as f:
                    writer.write(f)
            return send_file(part_path, as_attachment=True)
        elif tool == 'lock':
            password = options.get('password', '1234')
            writer = PdfWriter()
            reader = PdfReader(saved_files[0])
            for page in reader.pages:
                writer.add_page(page)
            writer.encrypt(password)
            with open(output_file, "wb") as f:
                writer.write(f)
        elif tool == 'convert':
            convert_to = options.get('convertTo', 'jpg')
            images = convert_from_path(saved_files[0])
            output_images = []
            for i, img in enumerate(images):
                img_path = os.path.join(DOWNLOAD_FOLDER, f"page_{i+1}.{convert_to}")
                img.save(img_path)
                output_images.append(img_path)
            return send_file(output_images[0], as_attachment=True)
        else:
            return "Tool not supported", 400
    except Exception as e:
        return f"Error: {e}", 500

    return send_file(output_file, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
