from flask import Flask, render_template, request, send_file, abort
import pandas as pd
from docxtpl import DocxTemplate
import zipfile
import os
import logging
from dotenv import load_dotenv
load_dotenv()

# Setup logging
logging.basicConfig(level=logging.INFO)

# Dynamic folder paths
BASE_DIR = os.path.abspath(os.path.dirname(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, os.getenv('UPLOAD_FOLDER', 'uploads'))
REPORT_FOLDER = os.path.join(BASE_DIR, os.getenv('REPORT_FOLDER', 'generated_reports'))
ZIP_NAME = os.getenv('ZIP_NAME', 'report_cards.zip')

# Ensure folders exist
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(REPORT_FOLDER, exist_ok=True)

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload_template', methods=['POST'])
def upload_template():
    try:
        template_file = request.files['template']
        template_path = os.path.join(UPLOAD_FOLDER, 'template.docx')
        template_file.save(template_path)

        doc = DocxTemplate(template_path)
        placeholders = doc.get_undeclared_template_variables()

        df = pd.DataFrame(columns=list(placeholders))
        excel_path = os.path.join(UPLOAD_FOLDER, 'custom_excel.xlsx')
        df.to_excel(excel_path, index=False)

        return send_file(excel_path, as_attachment=True)
    except Exception as e:
        logging.error(f"Error in upload_template: {e}")
        abort(500)

@app.route('/generate_reports', methods=['POST'])
def generate_reports():
    try:
        excel_file = request.files['excel']
        excel_path = os.path.join(UPLOAD_FOLDER, 'data.xlsx')
        excel_file.save(excel_path)

        template_path = os.path.join(UPLOAD_FOLDER, 'template.docx')
        df = pd.read_excel(excel_path)

        # Clear old reports
        for file in os.listdir(REPORT_FOLDER):
            os.remove(os.path.join(REPORT_FOLDER, file))

        # Generate new reports
        for _, row in df.iterrows():
            doc = DocxTemplate(template_path)
            context = row.to_dict()
            filename = f"{context.get('Name', 'Unnamed')}_report.docx"
            filepath = os.path.join(REPORT_FOLDER, filename)
            doc.render(context)
            doc.save(filepath)

        # Create ZIP
        zip_path = os.path.join(BASE_DIR, ZIP_NAME)
        with zipfile.ZipFile(zip_path, 'w') as zipf:
            for file in os.listdir(REPORT_FOLDER):
                zipf.write(os.path.join(REPORT_FOLDER, file), arcname=file)

        return send_file(zip_path, as_attachment=True)
    except Exception as e:
        logging.error(f"Error in generate_reports: {e}")
        abort(500)

if __name__ == '__main__':
    app.run(debug=False, host='0.0.0.0', port=int(os.getenv('PORT', 5000)))

