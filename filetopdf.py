from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import FileResponse
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

app = FastAPI()

UPLOAD_FOLDER = "uploads"
OUTPUT_FOLDER = "outputs"

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)


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
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
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
            # Preserve blank lines as vertical spacing
            content.append(Spacer(1, 0.2 * inch))

    pdf.build(content)


def convert_docx_to_pdf(input_path, output_path):
    # Use docx2pdf which leverages Microsoft Word for a full-fidelity conversion
    # (preserves formatting, tables, images, headings, etc.)
    docx2pdf_convert(input_path, output_path)


@app.post("/convert")
async def convert(file: UploadFile = File(...)):
    unique_name = str(uuid.uuid4()) + "_" + file.filename
    input_path = os.path.join(UPLOAD_FOLDER, unique_name)

    # Save uploaded file
    with open(input_path, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)

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
            raise HTTPException(status_code=400, detail=f"Unsupported file type: .{file_ext}")

    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Conversion failed: {str(e)}")

    finally:
        # Clean up the uploaded input file
        if os.path.exists(input_path):
            os.remove(input_path)

    return FileResponse(
        output_path,
        media_type="application/pdf",
        filename="output.pdf"
    )