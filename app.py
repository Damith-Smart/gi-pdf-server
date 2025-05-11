from flask import Flask, request, send_file
from openpyxl import load_workbook
import os

app = Flask(__name__)


@app.route("/generate-pdf", methods=["POST"])
def generate_pdf():
    data = request.json

    wb = load_workbook("gi_template.xlsx")
    ws = wb.active

    # ✅ Input Fields (let Excel do its calculations where formulas already exist)
    ws["E5"] = data.get("date", "")
    ws["E6"] = data.get("blast_in_charge", "")
    ws["E7"] = data.get("driller", "")
    ws["E9"] = data.get("time", "")
    ws["E10"] = data.get("material_type", "")
    ws["E11"] = data.get("location", "")
    ws["E12"] = data.get("no_of_holes", "")
    ws["E13"] = data.get("hole_depth", "")
    ws["E14"] = data.get("sub_holes", "")
    ws["E15"] = data.get("sub_depth", "")
    ws["E16"] = data.get("spacing", "")
    ws["E17"] = data.get("burden", "")
    ws["E18"] = data.get("density", "")

    # ✅ Watergel per hole (convert cartridges to weight)
    watergel_cartridges = int(data.get("watergel_per_hole", 12))
    ws["L9"] = watergel_cartridges * 0.125

    # ✅ Ammonium Nitrate per hole
    ws["L29"] = data.get("ammonium_per_hole", "")

    # ✅ ED Pattern (input only counts, not label text)
    ed_pattern = data.get("ed_pattern", [0] * 10)
    for i in range(10):
        ws[f"L{15 + i}"] = int(ed_pattern[i])  # just the integer

    # ✅ Save and Convert
    filled_path = "filled_gi_form.xlsx"
    pdf_path = "gi_form.pdf"
    wb.save(filled_path)
    os.system(
        f"libreoffice --headless --convert-to pdf {filled_path} --outdir .")

    if os.path.exists(pdf_path):
        return send_file(pdf_path, as_attachment=True)
    else:
        return {"error": "PDF generation failed"}, 500


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
