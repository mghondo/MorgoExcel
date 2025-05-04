import os
import re
import pdfplumber
from flask import Blueprint, request, jsonify
from werkzeug.utils import secure_filename

buildscan_bp = Blueprint('buildscan', __name__)

UPLOAD_FOLDER = 'BUILD-IN'
ALLOWED_EXTENSIONS = {'pdf'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def extract_manifest_data(pdf_path):
    data = {
        "originating_entity": None,
        "drivers": [],
        "products": []
    }

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()

            # Extract originating entity
            if not data["originating_entity"]:
                entity_match = re.search(r'Originating Entity\s*\|\s*(.+?)\s*\n', text)
                data["originating_entity"] = entity_match.group(1).strip() if entity_match else "Not found"

            # Extract drivers
            driver_matches = re.finditer(
                r'Name of Person Transporting \| (.+?) \| Employee ID of Driver \| (.+?)\n',
                text
            )
            for match in driver_matches:
                data["drivers"].append({
                    "name": match.group(1).strip(),
                    "employee_id": match.group(2).strip()
                })

            # Extract product data
            product_blocks = re.split(r'(\d+\. Package \| Shipped)', text)
            for block in product_blocks[1:]:
                product_data = extract_product_data(block)
                if product_data:
                    data["products"].append(product_data)

    return data

def extract_product_data(block):
    product = {
        "package_id": None,
        "m_number": None,
        "product_name": None,
        "type": "Not Specified",
        "thc": None,
        "weight": None,
        "supply_days": None
    }

    # Package ID
    package_id_match = re.search(r'(1A[\dA-Z]+)', block)
    if package_id_match:
        product["package_id"] = package_id_match.group(1)

    # M Number
    m_number_match = re.search(r'(M\d{8,})', block)
    if m_number_match:
        product["m_number"] = m_number_match.group(1)

    # Product name and type
    name_match = re.search(r'Item Name \| (.+?) \|', block)
    if name_match:
        name_parts = name_match.group(1).split('-')
        product["product_name"] = name_parts[-1].strip()
        product["type"] = detect_product_type(name_match.group(1))

    # THC, Weight, Supply
    details_match = re.search(
        r'THC: ([\d.]+)%?.*?Wgt: ([\d.]+ g).*?Supply: ([\d.]+) day\(s\)',
        block, re.DOTALL
    )
    if details_match:
        product.update({
            "thc": f"{details_match.group(1)}%",
            "weight": details_match.group(2),
            "supply_days": details_match.group(3)
        })

    return product

def detect_product_type(text):
    type_keywords = ["Indica", "Sativa", "Hybrid"]
    for keyword in type_keywords:
        if keyword in text:
            return keyword
    return "Not Specified"

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
            raw_data = extract_manifest_data(file_path)
            
            # Transform data for React compatibility
            formatted_data = {
                'company': raw_data['originating_entity'],
                'drivers': " / ".join([
                    f"{driver['name']} ({driver['employee_id']})" 
                    for driver in raw_data['drivers']
                ]),
                'items': [
                    {
                        'item_number': idx+1,
                        'package_id': product['package_id'],
                        'm_number': product['m_number'],
                        'name': product['product_name'],
                        'type': product['type'],
                        'strain': product['type'],  # Map type to strain
                        'days': product['supply_days'],
                        'weight': product['weight']
                    }
                    for idx, product in enumerate(raw_data['products'])
                ]
            }
            
            os.remove(file_path)
            return jsonify(formatted_data), 200

        except Exception as e:
            return jsonify({'error': str(e)}), 500

    return jsonify({'error': 'Invalid file type'}), 400
