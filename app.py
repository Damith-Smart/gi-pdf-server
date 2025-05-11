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

    # Fill Excel cells from request data
    ws["E5"] = data.get("date", "2025-05-10")
    ws["E6"] = data.get("blast_in_charge", "Millhas")
    ws["E7"] = data.get("driller", "Kaveesha")
    ws["E9"] = data.get("time", "10.00 a.m")
    ws["E10"] = data.get("material_type", "High Grade")
    ws["E111"] = data.get("location", "ISURU")

    ws["E12"] = data.get("no_of_holes", 10)
    ws["E13"] = data.get("hole_depth", 10.3)
    ws["E14"] = data.get("sub_holes", 10)
    ws["E15"] = data.get("sub_depth", 10.3)

    ws["E16"] = data.get("spacing", 2.8)
    ws["E17"] = data.get("burden", 2.5)
    ws["E18"] = data.get("density", 2.4)

    ws["L9"] = data.get("watergel_per_hole", "12 cartridges")
    ws["L29"] = data.get("ammonium_per_hole", "20 kg")

    # Handle ED Pattern List
    ed_pattern = data.get("ed_pattern", ["0"] * 10)
    for i in range(10):
        ws[f"L{15 + i}"] = f"ED number {i:02d}: {ed_pattern[i]}"

    # Save Excel file
    filled_path = "filled_gi_form.xlsx"
    pdf_path = "gi_form.pdf"
    wb.save(filled_path)

    # Convert to PDF using LibreOffice (must be installed in Render)
    os.system(
        f"libreoffice --headless --convert-to pdf {filled_path} --outdir .")

    # Return the file if generated
    if os.path.exists(pdf_path):
        return send_file(pdf_path, as_attachment=True)
    else:
        return {"error": "PDF generation failed"}, 500


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
