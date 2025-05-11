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

    # Fill Excel cells from request data (you only fill input cells)
    ws["E5"] = data.get("date", "2025-05-10")
    ws["E6"] = data.get("blast_in_charge", "Millhas")
    ws["E7"] = data.get("driller", "Kaveesha")
    ws["E9"] = data.get("time", "10.00 a.m")
    ws["E10"] = data.get("material_type", "High Grade")
    ws["E11"] = data.get("location", "ISURU")

    ws["E13"] = data.get("hole_depth", 10.3)
    ws["E15"] = data.get("sub_depth", 10.3)

    ws["E16"] = data.get("spacing", 2.8)
    ws["E17"] = data.get("burden", 2.5)
    ws["E18"] = data.get("density", 2.4)

    # Delay count per ED number (only int)
    ed_pattern = data.get("ed_pattern", ["0"] * 10)
    for i in range(10):
        try:
            count = int(ed_pattern[i])  # Force to int
        except:
            count = 0
        ws[f"L{15 + i}"] = count  # ED count goes to L15–L24

    # Watergel and Ammonium Nitrate (per hole only)
    ws["L9"] = data.get("watergel_per_hole", 0.125)  # cartridge weight
    ws["L29"] = data.get("ammonium_per_hole", 20)  # ammonium weight per hole

    # Save filled Excel file
    filled_path = "filled_gi_form.xlsx"
    wb.save(filled_path)

    # Just return Excel (skip PDF)
    return send_file(filled_path, as_attachment=True)


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
