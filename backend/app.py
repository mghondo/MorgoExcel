import os
import shutil
from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS
from werkzeug.utils import secure_filename
from datetime import datetime
from MorningRUN import process_excel
from WeeklyRUN import process_weekly_csv

app = Flask(__name__)
CORS(app)

MORNING_UPLOAD_FOLDER = 'MORNINGDROP'
MORNING_COMPLETE_FOLDER = 'MORNINGCOMPLETE'
WEEKLY_UPLOAD_FOLDER = 'WEEKLYDROP'
WEEKLY_COMPLETE_FOLDER = 'WEEKLYCOMPLETE'
ALLOWED_EXTENSIONS = {'xlsx', 'csv'}

app.config['MORNING_UPLOAD_FOLDER'] = MORNING_UPLOAD_FOLDER
app.config['MORNING_COMPLETE_FOLDER'] = MORNING_COMPLETE_FOLDER
app.config['WEEKLY_UPLOAD_FOLDER'] = WEEKLY_UPLOAD_FOLDER
app.config['WEEKLY_COMPLETE_FOLDER'] = WEEKLY_COMPLETE_FOLDER

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

# def clear_folders():
#     folders = [MORNING_UPLOAD_FOLDER, MORNING_COMPLETE_FOLDER, WEEKLY_UPLOAD_FOLDER, WEEKLY_COMPLETE_FOLDER]
#     for folder in folders:
#         if os.path.exists(folder):
#             for filename in os.listdir(folder):
#                 file_path = os.path.join(folder, filename)
#                 try:
#                     if os.path.isfile(file_path) or os.path.islink(file_path):
#                         os.unlink(file_path)
#                         print(f"Deleted file: {file_path}")
#                     elif os.path.isdir(file_path):
#                         shutil.rmtree(file_path)
#                         print(f"Deleted directory: {file_path}")
#                 except Exception as e:
#                     print(f'Failed to delete {file_path}. Reason: {e}')
#     print("All folders cleared successfully.")

@app.route('/upload/morning', methods=['POST'])
def upload_morning_file():
    if 'file' not in request.files:
        return jsonify({'error': 'No file part'}), 400
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'No selected file'}), 400
    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        file.save(os.path.join(app.config['MORNING_UPLOAD_FOLDER'], filename))
        input_path = os.path.join(app.config['MORNING_UPLOAD_FOLDER'], filename)
        output_filename = process_excel(input_path)
        print(f"Morning file processed: {output_filename}")
        return jsonify({'filename': output_filename}), 200

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
            output_file = process_weekly_csv()
            output_filename = os.path.basename(output_file)
            print(f"Weekly file processed: {output_filename}")
            return jsonify({'filename': output_filename}), 200
        except Exception as e:
            print(f"Error processing weekly file: {str(e)}")
            return jsonify({'error': str(e)}), 500

@app.route('/download/<folder>/<filename>', methods=['GET'])
def download_file(folder, filename):
    if folder == 'morning':
        return send_from_directory(app.config['MORNING_COMPLETE_FOLDER'], filename)
    elif folder == 'weekly':
        return send_from_directory(app.config['WEEKLY_COMPLETE_FOLDER'], filename)
    else:
        return jsonify({'error': 'Invalid folder'}), 400



if __name__ == '__main__':
    print("Starting Flask application...")
    os.makedirs(MORNING_UPLOAD_FOLDER, exist_ok=True)
    os.makedirs(MORNING_COMPLETE_FOLDER, exist_ok=True)
    os.makedirs(WEEKLY_UPLOAD_FOLDER, exist_ok=True)
    os.makedirs(WEEKLY_COMPLETE_FOLDER, exist_ok=True)
    # clear_folders()
    print("Folders created and cleared. Starting server...")
    app.run(debug=True, host='0.0.0.0', port=5000)
