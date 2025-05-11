from flask import Flask, request, send_file
from openpyxl import load_workbook
import os

app = Flask(__name__)


@app.route("/generate-pdf", methods=["POST"])
def generate_pdf():
    data = request.json
    wb = load_workbook("gi_template.xlsx")
    ws = wb.active

    # ✅ Fill input cells exactly
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

    # ✅ Explosives (input kg directly)
    ws["L9"] = data.get("watergel_per_hole", "")
    ws["L10"] = data.get("watergel_per_sub_hole", "")
    ws["L31"] = data.get("ammonium_per_hole", "")
    ws["L32"] = data.get("ammonium_per_sub_hole", "")

    # ✅ ED Pattern (delay counts only)
    ed_pattern = data.get("ed_pattern", [0] * 10)
    for i in range(10):
        try:
            ws[f"L{15 + i}"] = int(ed_pattern[i])
        except:
            ws[f"L{15 + i}"] = ""

    # ✅ Save and return .xlsx
    output_path = "filled_gi_form.xlsx"
    wb.save(output_path)
    return send_file(
        output_path,
        as_attachment=True,
        download_name="filled_gi_form.xlsx",
        mimetype=
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
