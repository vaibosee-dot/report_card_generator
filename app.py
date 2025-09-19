from flask import Flask, render_template, request, send_file
import pandas as pd
from docxtpl import DocxTemplate
import zipfile
import os

app = Flask(__name__)

# Ensure folders exist
os.makedirs('uploads', exist_ok=True)
os.makedirs('generated_reports', exist_ok=True)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload_template', methods=['POST'])
def upload_template():
    template_file = request.files['template']
    template_path = os.path.join("uploads", "template.docx")
    template_file.save(template_path)

    # Extract placeholders from template
    doc = DocxTemplate(template_path)
    placeholders = doc.get_undeclared_template_variables()

    # Create Excel with those placeholders as columns
    df = pd.DataFrame(columns=list(placeholders))
    excel_path = os.path.join("uploads", "custom_excel.xlsx")
    df.to_excel(excel_path, index=False)

    return send_file(excel_path, as_attachment=True)

@app.route('/generate_reports', methods=['POST'])
def generate_reports():
    excel_file = request.files['excel']
    excel_path = os.path.join("uploads", "data.xlsx")
    excel_file.save(excel_path)

    template_path = os.path.join("uploads", "template.docx")

    # Read Excel data
    df = pd.read_excel(excel_path)

    # Clear old reports
    for file in os.listdir("generated_reports"):
        os.remove(os.path.join("generated_reports", file))

    # Generate report cards
    for _, row in df.iterrows():
        doc = DocxTemplate(template_path)
        context = row.to_dict()
        filename = f"{row['Name']}_report.docx"
        filepath = os.path.join("generated_reports", filename)
        doc.render(context)
        doc.save(filepath)

    # Create ZIP file
    zip_path = "report_cards.zip"
    with zipfile.ZipFile(zip_path, 'w') as zipf:
        for file in os.listdir("generated_reports"):
            zipf.write(os.path.join("generated_reports", file), arcname=file)

    return send_file(zip_path, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)

