from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
import pandas as pd
import os
from werkzeug.utils import secure_filename
from sample import generate_test_cases
from createcase import TestCaseManager
import json
from datetime import datetime

app = Flask(__name__)
CORS(app)  # Enable CORS for frontend communication

# Configuration
UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'outputs'
TEMPLATE_FILE = 'sample_input.xlsx'
ALLOWED_EXTENSIONS = {'xlsx'}

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['OUTPUT_FOLDER'] = OUTPUT_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size


def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


@app.route('/api/health', methods=['GET'])
def health_check():
    """Health check endpoint"""
    return jsonify({'status': 'ok', 'message': 'Backend is running'})


@app.route('/api/download-template', methods=['GET'])
def download_template():
    """Download Excel template"""
    try:
        if os.path.exists(TEMPLATE_FILE):
            return send_file(
                TEMPLATE_FILE,
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                as_attachment=True,
                download_name='test_case_template.xlsx'
            )
        else:
            return jsonify({'error': 'Template file not found'}), 404
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/upload', methods=['POST'])
def upload_file():
    """Handle file upload"""
    if 'file' not in request.files:
        return jsonify({'error': 'No file provided'}), 400

    file = request.files['file']

    if file.filename == '':
        return jsonify({'error': 'No file selected'}), 400

    if not allowed_file(file.filename):
        return jsonify({'error': 'Invalid file type. Only .xlsx files are allowed'}), 400

    try:
        filename = secure_filename(file.filename)
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        unique_filename = f"{timestamp}_{filename}"
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], unique_filename)
        file.save(filepath)

        return jsonify({
            'success': True,
            'filename': unique_filename,
            'message': 'File uploaded successfully'
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/generate', methods=['POST'])
def generate():
    """Generate test cases"""
    try:
        data = request.json
        filename = data.get('filename')
        ado_config = data.get('adoConfig', {})
        upload_to_ado = data.get('uploadToAdo')

        if not filename:
            return jsonify({'error': 'No filename provided'}), 400

        input_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)

        if not os.path.exists(input_path):
            return jsonify({'error': 'Uploaded file not found'}), 404

        # Prepare output file
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        output_filename = f"generated_test_cases_{timestamp}.xlsx"
        output_path = os.path.join(app.config['OUTPUT_FOLDER'], output_filename)

        # Initialize TestCaseManager if ADO upload is enabled
        mgr = None
        suite_name = None

        if upload_to_ado:
            try:
                mgr = TestCaseManager(
                    org=ado_config.get('organization'),
                    proj=ado_config.get('project'),
                    pat=ado_config.get('pat'),
                    plan_name=ado_config.get('planName')
                )
                suite_name = ado_config.get('suiteName')
            except Exception as e:
                return jsonify({'error': f'ADO connection failed: {str(e)}'}), 500

        # Generate test cases
        df_output = generate_test_cases(
            input_file=input_path,
            output_file=output_path,
            mgr=mgr,
            suite_name=suite_name
        )

        # Convert DataFrame to JSON for response
        test_cases = df_output.to_dict('records')

        return jsonify({
            'success': True,
            'testCases': test_cases,
            'outputFilename': output_filename,
            'message': f'Generated {len(test_cases)} test cases successfully',
            'uploadedToAdo': upload_to_ado
        })

    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/download-results/<filename>', methods=['GET'])
def download_results(filename):
    """Download generated test cases"""
    try:
        filepath = os.path.join(app.config['OUTPUT_FOLDER'], filename)

        if not os.path.exists(filepath):
            return jsonify({'error': 'File not found'}), 404

        return send_file(
            filepath,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name='generated_test_cases.xlsx'
        )
    except Exception as e:
        return jsonify({'error': str(e)}), 500



if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)