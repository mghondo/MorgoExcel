from flask import Flask, request, jsonify, send_from_directory
from werkzeug.utils import secure_filename
import os
import datetime
from flask_cors import CORS

# Import custom modules
from metric import process_metric_file
from dutchie import process_dutchie_file, clear_output_directory
import MorningRUN
import WeeklyRUN
from order import process_order_file
from email_sender import send_email
from buildscan import buildscan_bp, scan_pdf

app = Flask(__name__)
# CORS(app)

# CORS(app, resources={r"/*": {"origins": ["https://morgotools.com", "http://localhost:3000"]}})
CORS(app, resources={r"/*": {"origins": ["https://morgotools.com", "http://localhost:3000", "https://www.morgotools.com"]}})

# Register the buildscan blueprint
app.register_blueprint(buildscan_bp)

# File upload configurations
UPLOAD_FOLDER = 'uploads'
MORNING_UPLOAD_FOLDER = 'MORNINGDROP'
WEEKLY_UPLOAD_FOLDER = 'WEEKLYDROP'
METRIC_UPLOAD_FOLDER = 'METRIC-IN'
DUTCHIE_UPLOAD_FOLDER = 'DUTCHIE-IN'
ORDER_UPLOAD_FOLDER = 'ORDER-IN'
BUILD_IN_FOLDER = 'BUILD-IN'
MORNING_COMPLETE_FOLDER = 'MORNINGCOMPLETE'
WEEKLY_COMPLETE_FOLDER = 'WEEKLYCOMPLETE'
METRIC_COMPLETE_FOLDER = 'METRIC-OUT'
DUTCHIE_COMPLETE_FOLDER = 'DUTCHIE-OUT'
ORDER_COMPLETE_FOLDER = 'ORDER-OUT'

ALLOWED_EXTENSIONS = {'csv', 'xlsx', 'pdf'}

# Configure upload folders in Flask app
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['MORNING_UPLOAD_FOLDER'] = MORNING_UPLOAD_FOLDER
app.config['WEEKLY_UPLOAD_FOLDER'] = WEEKLY_UPLOAD_FOLDER
app.config['METRIC_UPLOAD_FOLDER'] = METRIC_UPLOAD_FOLDER
app.config['DUTCHIE_UPLOAD_FOLDER'] = DUTCHIE_UPLOAD_FOLDER
app.config['ORDER_UPLOAD_FOLDER'] = ORDER_UPLOAD_FOLDER
app.config['BUILD_IN_FOLDER'] = BUILD_IN_FOLDER
app.config['MORNING_COMPLETE_FOLDER'] = MORNING_COMPLETE_FOLDER
app.config['WEEKLY_COMPLETE_FOLDER'] = WEEKLY_COMPLETE_FOLDER
app.config['METRIC_COMPLETE_FOLDER'] = METRIC_COMPLETE_FOLDER
app.config['DUTCHIE_COMPLETE_FOLDER'] = DUTCHIE_COMPLETE_FOLDER
app.config['ORDER_COMPLETE_FOLDER'] = ORDER_COMPLETE_FOLDER

# Helper function to check allowed file types
def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

# ------------------- Existing Routes -------------------

@app.route('/send-email', methods=['POST'])
def handle_email():
    try:
        data = request.json
        result = send_email(
            selected_items=data.get('selectedItems', {}),
            email_type=data.get('emailType', 'low_stock'),
            location=data.get('location', 'Unknown Location')
        )
        
        if result['status'] == 'success':
            return jsonify({"message": "Email sent successfully!"}), 200
        else:
            return jsonify({"error": result['message']}), 500
            
    except Exception as e:
        return jsonify({"error": str(e)}), 500

# Morning file routes (preserved)
@app.route('/upload/morning', methods=['POST'])
def upload_morning_file():
    if 'file' not in request.files:
        return jsonify({'error': 'No file part'}), 400

    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'No selected file'}), 400

    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        file_path = os.path.join(app.config['MORNING_UPLOAD_FOLDER'], filename)
        file.save(file_path)

        try:
            processed_filename = MorningRUN.process_morning_file(file_path)
            return jsonify({'filename': processed_filename}), 200
        except Exception as e:
            return jsonify({'error': str(e)}), 500

@app.route('/download/morning/<filename>', methods=['GET'])
def download_morning_file(filename):
    try:
        return send_from_directory(app.config['MORNING_COMPLETE_FOLDER'], filename, as_attachment=True)
    except Exception as e:
        return jsonify({'error': str(e)}), 500

# Weekly file routes (preserved)
@app.route('/upload/weekly', methods=['POST'])
def upload_weekly_file():
    if 'file' not in request.files:
        return jsonify({'error': 'No file part'}), 400

    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'No selected file'}), 400

    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        file_path = os.path.join(app.config['WEEKLY_UPLOAD_FOLDER'], filename)
        file.save(file_path)

        try:
            processed_filename = WeeklyRUN.process_weekly_file(file_path)
            return jsonify({'filename': processed_filename}), 200
        except Exception as e:
            return jsonify({'error': str(e)}), 500

@app.route('/download/weekly/<filename>', methods=['GET'])
def download_weekly_file(filename):
    try:
        return send_from_directory(app.config['WEEKLY_COMPLETE_FOLDER'], filename, as_attachment=True)
    except Exception as e:
        return jsonify({'error': str(e)}), 500

# Metric file routes (preserved)
@app.route('/upload/metric', methods=['POST'])
def upload_metric_file():
    if 'file' not in request.files:
        return jsonify({'error': 'No file part'}), 400

    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'No selected file'}), 400

    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        input_path = os.path.join(app.config['METRIC_UPLOAD_FOLDER'], filename)
        file.save(input_path)

        try:
            output_filename = f'METRIC-{filename[:-4]}.xlsx'
            output_path = os.path.join(app.config['METRIC_COMPLETE_FOLDER'], output_filename)
            process_metric_file(input_path, output_path)
            return jsonify({'filename': output_filename}), 200
        except Exception as e:
            return jsonify({'error': str(e)}), 500

@app.route('/download/metric/<filename>', methods=['GET'])
def download_metric_file(filename):
    try:
        return send_from_directory(app.config['METRIC_COMPLETE_FOLDER'], filename, as_attachment=True)
    except Exception as e:
        return jsonify({'error': str(e)}), 500

# Dutchie file routes (preserved)
@app.route('/upload/dutchie', methods=['POST'])
def upload_dutchie_file():
    if 'file' not in request.files:
        return jsonify({'error': 'No file part'}), 400

    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'No selected file'}), 400

    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        input_path = os.path.join(app.config['DUTCHIE_UPLOAD_FOLDER'], filename)
        file.save(input_path)

        try:
            clear_output_directory(app.config['DUTCHIE_COMPLETE_FOLDER'])
            process_dutchie_file(input_path, None)
            current_date = datetime.date.today().strftime("%m-%d-%Y")
            marengo_filename = f'Marengo-Dutchie-{current_date}.xlsx'
            columbus_filename = f'Columbus-Dutchie-{current_date}.xlsx'
            output_filename = marengo_filename if os.path.exists(os.path.join(app.config['DUTCHIE_COMPLETE_FOLDER'], marengo_filename)) else columbus_filename
            return jsonify({'filename': output_filename}), 200
        except Exception as e:
            return jsonify({'error': str(e)}), 500

@app.route('/download/dutchie/<filename>', methods=['GET'])
def download_dutchie_file(filename):
    try:
        return send_from_directory(app.config['DUTCHIE_COMPLETE_FOLDER'], filename, as_attachment=True)
    except Exception as e:
        return jsonify({'error': str(e)}), 500

# ------------------- New Ordering Routes -------------------

@app.route('/process-order', methods=['POST'])
def process_order():
    if 'file' not in request.files or 'numOfDays' not in request.form:
        return jsonify({'error': 'File and number of days are required'}), 400

    file = request.files['file']
    num_of_days_str = request.form.get('numOfDays')

    if not num_of_days_str.isdigit() or int(num_of_days_str) <= 0:
        return jsonify({'error': 'Invalid number of days'}), 400

    num_of_days = int(num_of_days_str)

    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        input_path = os.path.join(app.config['ORDER_UPLOAD_FOLDER'], filename)
        file.save(input_path)

        try:
            vendor_data, location = process_order_file(input_path, num_of_days)
            output_filename = f'ORDER-{filename[:-5]}-{datetime.date.today().strftime("%m-%d-%Y")}.xlsx'
            output_path = os.path.join(app.config['ORDER_COMPLETE_FOLDER'], output_filename)
            # Save processed data to output_path (if needed by the frontend later).
            return jsonify({'vendor_data': vendor_data, 'location': location}), 200
        except Exception as e:
            return jsonify({'error': str(e)}), 500

@app.route('/download/order/<filename>', methods=['GET'])
def download_order_file(filename):
    try:
        return send_from_directory(app.config['ORDER_COMPLETE_FOLDER'], filename, as_attachment=True)
    except Exception as e:
        return jsonify({'error': str(e)}), 500
    
@app.route('/health', methods=['GET'])
def health_check():
    return {"status": "healthy"}, 200




# ------------------- Main Application Entry -------------------

if __name__ == '__main__':
    # Ensure all required directories exist before running the app.
    os.makedirs(UPLOAD_FOLDER, exist_ok=True)
    os.makedirs(MORNING_UPLOAD_FOLDER, exist_ok=True)
    os.makedirs(WEEKLY_UPLOAD_FOLDER, exist_ok=True)
    os.makedirs(METRIC_UPLOAD_FOLDER, exist_ok=True)
    os.makedirs(DUTCHIE_UPLOAD_FOLDER, exist_ok=True)
    os.makedirs(ORDER_UPLOAD_FOLDER, exist_ok=True)
    os.makedirs(BUILD_IN_FOLDER, exist_ok=True)
    os.makedirs(MORNING_COMPLETE_FOLDER, exist_ok=True)
    os.makedirs(WEEKLY_COMPLETE_FOLDER, exist_ok=True)
    os.makedirs(METRIC_COMPLETE_FOLDER, exist_ok=True)
    os.makedirs(DUTCHIE_COMPLETE_FOLDER, exist_ok=True)
    os.makedirs(ORDER_COMPLETE_FOLDER, exist_ok=True)

    app.run(debug=True)
