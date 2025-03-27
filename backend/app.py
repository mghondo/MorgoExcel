from flask import Flask, request, jsonify, send_from_directory
from werkzeug.utils import secure_filename
import os
from metric import process_metric_file
from dutchie import process_dutchie_file, clear_output_directory
import datetime
from flask_cors import CORS
import MorningRUN
import WeeklyRUN
from order import process_order_file
from email_sender import send_email  # New import for email functionality

app = Flask(__name__)
CORS(app)

current_date = datetime.date.today().strftime("%m-%d-%Y")

# File upload configurations
UPLOAD_FOLDER = 'uploads'
MORNING_UPLOAD_FOLDER = 'MORNINGDROP'
WEEKLY_UPLOAD_FOLDER = 'WEEKLYDROP'
METRIC_UPLOAD_FOLDER = 'METRIC-IN'
DUTCHIE_UPLOAD_FOLDER = 'DUTCHIE-IN'
ORDER_UPLOAD_FOLDER = 'ORDER-IN'
MORNING_COMPLETE_FOLDER = 'MORNINGCOMPLETE'
WEEKLY_COMPLETE_FOLDER = 'WEEKLYCOMPLETE'
METRIC_COMPLETE_FOLDER = 'METRIC-OUT'
DUTCHIE_COMPLETE_FOLDER = 'DUTCHIE-OUT'
ALLOWED_EXTENSIONS = {'csv', 'xlsx'}

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['MORNING_UPLOAD_FOLDER'] = MORNING_UPLOAD_FOLDER
app.config['WEEKLY_UPLOAD_FOLDER'] = WEEKLY_UPLOAD_FOLDER
app.config['METRIC_UPLOAD_FOLDER'] = METRIC_UPLOAD_FOLDER
app.config['DUTCHIE_UPLOAD_FOLDER'] = DUTCHIE_UPLOAD_FOLDER
app.config['ORDER_UPLOAD_FOLDER'] = ORDER_UPLOAD_FOLDER
app.config['MORNING_COMPLETE_FOLDER'] = MORNING_COMPLETE_FOLDER
app.config['WEEKLY_COMPLETE_FOLDER'] = WEEKLY_COMPLETE_FOLDER
app.config['METRIC_COMPLETE_FOLDER'] = METRIC_COMPLETE_FOLDER
app.config['DUTCHIE_COMPLETE_FOLDER'] = DUTCHIE_COMPLETE_FOLDER

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

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

@app.route('/download/<folder>/<filename>', methods=['GET'])
def download_file(folder, filename):
    if folder == 'morning':
        return send_from_directory(app.config['MORNING_COMPLETE_FOLDER'], filename)
    elif folder == 'weekly':
        return send_from_directory(app.config['WEEKLY_COMPLETE_FOLDER'], filename)
    elif folder == 'metric':
        return send_from_directory(app.config['METRIC_COMPLETE_FOLDER'], filename)
    elif folder == 'dutchie':
        return send_from_directory(app.config['DUTCHIE_COMPLETE_FOLDER'], filename)
    else:
        return jsonify({'error': 'Invalid folder'}), 400

@app.route('/process-order', methods=['POST'])
def handle_order_processing():
    if 'file' not in request.files:
        return jsonify({'error': 'No file part'}), 400
    file = request.files['file']
    num_days = request.form.get('numOfDays', type=int)
    if not num_days or num_days < 1 or num_days > 999:
        return jsonify({'error': 'Invalid number of days (1-999 required)'}), 400
    if file.filename == '':
        return jsonify({'error': 'No selected file'}), 400
    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        file_path = os.path.join(app.config['ORDER_UPLOAD_FOLDER'], filename)
        file.save(file_path)
        try:
            # processed_data = process_order_file(file_path, num_days)
            # return jsonify(processed_data), 200
            processed_data, location = process_order_file(file_path, num_days)
            return jsonify({'vendor_data': processed_data, 'location': location}), 200

        except Exception as e:
            return jsonify({'error': str(e)}), 500
    return jsonify({'error': 'Invalid file type'}), 400

# New route for sending emails
@app.route('/send-email', methods=['POST'])
def handle_send_email():
    data = request.json
    selected_items = data.get('selectedItems', {})
    email_type = data.get('emailType', '')
    location = data.get('location', '')
    try:
        result = send_email(selected_items, email_type, location)
        return jsonify(result), 200
    except Exception as e:
        return jsonify({'error': str(e)}), 500



if __name__ == '__main__':
    os.makedirs(UPLOAD_FOLDER, exist_ok=True)
    os.makedirs(MORNING_UPLOAD_FOLDER, exist_ok=True)
    os.makedirs(WEEKLY_UPLOAD_FOLDER, exist_ok=True)
    os.makedirs(METRIC_UPLOAD_FOLDER, exist_ok=True)
    os.makedirs(DUTCHIE_UPLOAD_FOLDER, exist_ok=True)
    os.makedirs(ORDER_UPLOAD_FOLDER, exist_ok=True)
    os.makedirs(MORNING_COMPLETE_FOLDER, exist_ok=True)
    os.makedirs(WEEKLY_COMPLETE_FOLDER, exist_ok=True)
    os.makedirs(METRIC_COMPLETE_FOLDER, exist_ok=True)
    os.makedirs(DUTCHIE_COMPLETE_FOLDER, exist_ok=True)
    app.run(debug=True)
