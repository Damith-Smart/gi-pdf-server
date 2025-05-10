from flask import Flask, request, send_file
from openpyxl import load_workbook
import os

app = Flask(__name__)

@app.route("/generate-pdf", methods=["POST"])
def generate_pdf():
    data = request.json

    # Load the Excel template
    wb = load_workbook("gi_template.xlsx")
    ws = wb.active

    # Example: Fill in a few fields (you can customize more later)
    ws["E4"] = data.get("date", "2025-05-10")
    ws["E5"] = data.get("blast_in_charge", "Millhas")
    ws["E6"] = data.get("driller", "Kaveesha")
    ws["E11"] = data.get("no_of_holes", 10)
    ws["E12"] = data.get("hole_depth", 10.3)

    filled_path = "filled_gi_form.xlsx"
    pdf_path = "gi_form.pdf"
    wb.save(filled_path)

    # Convert Excel to PDF using LibreOffice headless (assumes libreoffice is installed)
    os.system(f"libreoffice --headless --convert-to pdf {filled_path} --outdir .")

    if not os.path.exists(pdf_path):
        return {"error": "PDF generation failed"}, 500

    return send_file(pdf_path, as_attachment=True)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
