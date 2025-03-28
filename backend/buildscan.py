import os
from flask import Blueprint, request, jsonify
from werkzeug.utils import secure_filename
import PyPDF2
import re

buildscan_bp = Blueprint('buildscan', __name__)

UPLOAD_FOLDER = 'BUILD-IN'
ALLOWED_EXTENSIONS = {'pdf'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def extract_outbound_transporter(text):
    # Regex to extract 'Outbound Transporter'
    pattern = r'1\.\s*Outbound\s*Transporter\s*(.*?)\n'
    match = re.search(pattern, text)
    return match.group(1).strip() if match else None

@buildscan_bp.route('/scan-pdf', methods=['POST'])
def scan_pdf():
    if 'file' not in request.files:
        return jsonify({'error': 'No file part'}), 400

    file = request.files['file']

    if file.filename == '':
        return jsonify({'error': 'No selected file'}), 400

    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        file_path = os.path.join(UPLOAD_FOLDER, filename)
        file.save(file_path)

        try:
            with open(file_path, 'rb') as pdf_file:
                pdf_reader = PyPDF2.PdfReader(pdf_file)
                num_pages = len(pdf_reader.pages)
                text = ""
                for page in pdf_reader.pages:
                    text += page.extract_text()

            outbound_transporter = extract_outbound_transporter(text)

            os.remove(file_path)  # Remove the file after processing

            return jsonify({
                'results': {
                    'num_pages': num_pages,
                    'text_preview': text[:2500] + '...' if len(text) > 2500 else text,
                    'outbound_transporter': outbound_transporter
                }
            }), 200

        except Exception as e:
            return jsonify({'error': str(e)}), 2500

    return jsonify({'error': 'Invalid file type'}), 400
