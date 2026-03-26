from flask import Flask, request, send_file, render_template, jsonify
import pandas as pd
from reportlab.platypus import SimpleDocTemplate, Table, Paragraph, Image, TableStyle, Spacer
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, landscape
from reportlab.lib.units import inch
import shutil
import os
import uuid
from docx2pdf import convert as docx2pdf_convert

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
OUTPUT_FOLDER = "outputs"

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

ALLOWED_EXTENSIONS = {"xlsx", "xls", "png", "jpg", "jpeg", "txt", "docx"}


def allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[-1].lower() in ALLOWED_EXTENSIONS


def convert_excel_to_pdf(input_path, output_path):
    df = pd.read_excel(input_path)
    data = [df.columns.tolist()] + df.values.tolist()
    styles = getSampleStyleSheet()
    wrapped_data = []
    for row in data:
        wrapped_row = [Paragraph(str(cell), styles["Normal"]) for cell in row]
        wrapped_data.append(wrapped_row)

    pdf = SimpleDocTemplate(output_path, pagesize=landscape(letter))
    page_width = landscape(letter)[0]
    num_cols = len(data[0])
    col_width = page_width / num_cols
    col_widths = [col_width] * num_cols

    table = Table(wrapped_data, colWidths=col_widths)
    style = TableStyle([
        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 0), (-1, -1), 8),
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#6C63FF')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#CCCCCC')),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#F5F5FF')]),
        ('PADDING', (0, 0), (-1, -1), 6),
    ])
    table.setStyle(style)
    pdf.build([table])


def convert_image_to_pdf(input_path, output_path):
    pdf = SimpleDocTemplate(output_path, pagesize=letter)
    page_width, page_height = letter

    img = Image(input_path)
    img_width = img.imageWidth
    img_height = img.imageHeight
    scale = min(
        (page_width - 100) / img_width,
        (page_height - 100) / img_height
    )
    img.drawWidth = img_width * scale
    img.drawHeight = img_height * scale
    img.hAlign = 'CENTER'

    pdf.build([img])


def convert_text_to_pdf(input_path, output_path):
    styles = getSampleStyleSheet()

    with open(input_path, "r", encoding="utf-8") as f:
        lines = f.readlines()

    pdf = SimpleDocTemplate(output_path, pagesize=letter,
                            leftMargin=inch, rightMargin=inch,
                            topMargin=inch, bottomMargin=inch)
    content = []
    for line in lines:
        stripped = line.strip()
        if stripped:
            content.append(Paragraph(stripped, styles["Normal"]))
        else:
            content.append(Spacer(1, 0.2 * inch))

    pdf.build(content)


def convert_docx_to_pdf(input_path, output_path):
    docx2pdf_convert(input_path, output_path)


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/convert", methods=["POST"])
def convert():
    if "file" not in request.files:
        return jsonify({"error": "No file provided"}), 400

    file = request.files["file"]

    if file.filename == "":
        return jsonify({"error": "No file selected"}), 400

    if not allowed_file(file.filename):
        return jsonify({"error": f"Unsupported file type. Allowed: xlsx, xls, png, jpg, jpeg, txt, docx"}), 400

    unique_name = str(uuid.uuid4()) + "_" + file.filename
    input_path = os.path.join(UPLOAD_FOLDER, unique_name)
    file.save(input_path)

    file_ext = file.filename.rsplit(".", 1)[-1].lower()
    output_filename = unique_name.rsplit(".", 1)[0] + ".pdf"
    output_path = os.path.join(OUTPUT_FOLDER, output_filename)

    try:
        if file_ext in ["xlsx", "xls"]:
            convert_excel_to_pdf(input_path, output_path)
        elif file_ext in ["png", "jpg", "jpeg"]:
            convert_image_to_pdf(input_path, output_path)
        elif file_ext == "txt":
            convert_text_to_pdf(input_path, output_path)
        elif file_ext == "docx":
            convert_docx_to_pdf(input_path, output_path)
        else:
            return jsonify({"error": "Unsupported file type"}), 400

    except Exception as e:
        return jsonify({"error": f"Conversion failed: {str(e)}"}), 500

    finally:
        if os.path.exists(input_path):
            os.remove(input_path)

    return send_file(
        output_path,
        mimetype="application/pdf",
        as_attachment=True,
        download_name="converted_output.pdf"
    )


if __name__ == "__main__":
    app.run(debug=True, port=5000)
