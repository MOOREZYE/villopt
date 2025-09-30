from flask import Flask, request, render_template, send_from_directory, redirect, url_for, flash
import os
import pandas as pd
from jinja2 import Environment, FileSystemLoader
from reportlab.lib.pagesizes import LETTER
from reportlab.pdfgen import canvas
from reportlab.lib import colors
from reportlab.lib.units import inch
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import Table, TableStyle, Paragraph
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.utils import ImageReader
from reportlab.lib.fonts import addMapping
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics

app = Flask(__name__)
app.secret_key = 'supersecretkey'
UPLOAD_FOLDER = 'uploads'
INVOICE_FOLDER = 'invoices'
ALLOWED_EXTENSIONS = {'xlsx'}

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(INVOICE_FOLDER, exist_ok=True)

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        if 'file' not in request.files:
            flash('No file part')
            return redirect(request.url)
        file = request.files['file']
        if file.filename == '':
            flash('No selected file')
            return redirect(request.url)
        if file and allowed_file(file.filename):
            filepath = os.path.join(UPLOAD_FOLDER, file.filename)
            file.save(filepath)
            generate_invoices(filepath)
            flash('Invoices generated successfully!')
            return redirect(url_for('list_invoices'))
    return render_template('upload.html')

@app.route('/invoices')
def list_invoices():
    files = os.listdir(INVOICE_FOLDER)
    return render_template('invoices.html', files=files)

@app.route('/invoices/<filename>')
def download_invoice(filename):
    return send_from_directory(INVOICE_FOLDER, filename)

@app.route('/invoices/clear', methods=['POST'])
def clear_invoices():
    for filename in os.listdir(INVOICE_FOLDER):
        file_path = os.path.join(INVOICE_FOLDER, filename)
        if os.path.isfile(file_path):
            os.remove(file_path)
    flash('All invoices have been deleted.')
    return redirect(url_for('list_invoices'))

def generate_invoices(filepath):
    df = pd.read_excel(filepath)
    for _, row in df.iterrows():
        pdf_path = os.path.join(INVOICE_FOLDER, f"invoice_{row['Invoice Number']}.pdf")
        c = canvas.Canvas(pdf_path, pagesize=LETTER)
        width, height = LETTER
        margin = 40
        green = colors.HexColor('#4B8B3B')
        grey = colors.HexColor('#E5E5E5')
        # Add logo
        logo_path = '/Users/matthewmoore/Documents/VillOpt/assets/VO-page-001.jpg'
        if os.path.exists(logo_path):
            logo = ImageReader(logo_path)
            c.drawImage(logo, margin, height - margin - 60, width=80, height=50, mask='auto')
        
        # Header: Company Name (moved to the right of logo)
        c.setFont("Helvetica-Bold", 36)
        c.setFillColor(green)
        c.drawString(margin + 100, height - margin - 10, "The Village Optician Ltd")
        c.setFillColor(colors.black)
        c.setFont("Helvetica", 10)
        c.drawString(margin + 100, height - margin - 30, "MICHAEL GREENBERG, BSc (Hons), MCOptom.")
        c.setFont("Helvetica-Oblique", 9)
        c.drawString(margin + 100, height - margin - 45, "OPHTHALMIC OPTICIAN & CONTACT LENS PRACTITIONER")
        # Contact Info (no rectangles)
        c.setFont("Helvetica", 8)
        c.drawString(margin + 2, height - margin - 56, "470 BURY NEW ROAD, PRESTWICH")
        c.drawString(margin + 202, height - margin - 56, "MANCHESTER, M25 1AX")
        c.drawString(width - margin - 178, height - margin - 56, "Telephone 0161 773 0069")
        c.drawString(width - margin - 178, height - margin - 76, "Fax 0161 773 0170")
        # Email and website (no rectangle)
        c.setFont("Helvetica", 8)
        c.drawString(width/2 - 85, height - margin - 90, "Email: reception@thevillageopticianltd.co.uk")
        c.setFillColor(green)
        c.drawString(width/2 - 85, height - margin - 100, "www.thevillageopticianltd.co.uk")
        c.setFillColor(colors.black)
        # INVOICE label (centered, above details)
        c.setFont("Helvetica-Bold", 28)
        c.setFillColor(green)
        c.drawCentredString(width/2, height - margin - 150, "INVOICE")
        c.setFillColor(colors.black)
        # Invoice Details Box (above table)
        details_top = height - margin - 200
        box_height = 130  # Increased height to fit all fields
        c.rect(margin, details_top - box_height, 350, box_height, stroke=1, fill=0)
        c.setFont("Helvetica", 12)
        c.drawString(margin + 10, details_top - 25, f"INVOICE:   {row['Invoice Number']}")
        c.drawString(margin + 10, details_top - 40, f"DATE:      {row['Date']}")
        c.drawString(margin + 10, details_top - 55, f"SUPPLIER:  {row['Supplier']}")
        c.drawString(margin + 10, details_top - 70, f"ORDER NO:  {row['Order Number']}")
        c.setFont("Helvetica", 10)
        # Hardcoded address, start lower
        address_lines = ["Vista", "18 Eli Hacohen Street", "Jerusalem", "9551120"]
        for i, line in enumerate(address_lines):
            c.drawString(margin + 10, details_top - 85 - (i * 12), line)
        # Table for items (below details)
        table_top = details_top - box_height - 20
        table_data = [["Quantity", "Description", "Amount"]]
        table_data.append([str(row['Quantity']), str(row['Product']), str(row['Amount'])])
        for _ in range(10):
            table_data.append(["", "", ""])
        table_data.append(["", "TOTAL", str(row['Amount'])])
        table = Table(table_data, colWidths=[80, 300, 100], rowHeights=22)
        table.setStyle(TableStyle([
            ("BACKGROUND", (0,0), (-1,0), grey),
            ("TEXTCOLOR", (0,0), (-1,0), colors.black),
            ("ALIGN", (0,0), (-1,0), "CENTER"),
            ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
            ("FONTSIZE", (0,0), (-1,0), 12),
            ("FONTSIZE", (0,1), (-1,-1), 10),
            ("BOX", (0,0), (-1,-1), 1, colors.black),
            ("INNERGRID", (0,0), (-1,-1), 0.5, colors.black),
            ("LINEBELOW", (0,0), (-1,0), 1, colors.black),
            ("BACKGROUND", (0,-1), (-1,-1), colors.white),
            ("FONTNAME", (1,-1), (1,-1), "Helvetica-Bold"),
            ("FONTNAME", (2,-1), (2,-1), "Helvetica-Bold"),
            ("ALIGN", (2,-1), (2,-1), "RIGHT"),
        ]))
        table.wrapOn(c, width, height)
        table.drawOn(c, margin, table_top - (len(table_data) * 22))
        c.save()

if __name__ == '__main__':
    app.run(debug=True) 