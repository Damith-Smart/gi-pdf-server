from flask import Flask, request, send_file
from openpyxl import load_workbook
import os

app = Flask(__name__)

@app.route("/generate-pdf", methods=["POST"])
def generate_pdf():
    try:
        data = request.json
        wb = load_workbook("gi_template.xlsx")
        ws = wb.active

        # 🟢 Basic Info
        ws["E5"] = data.get("date", "")
        ws["E6"] = data.get("blast_in_charge", "")
        ws["E7"] = data.get("driller", "")
        ws["E9"] = data.get("time", "")
        ws["E10"] = data.get("material_type", "")
        ws["E11"] = data.get("location", "")

        # 🟢 Blasting Design
        ws["E12"] = data.get("no_of_holes", 0)
        ws["E13"] = data.get("hole_depth", 0.0)
        ws["E14"] = data.get("sub_holes", 0)
        ws["E15"] = data.get("sub_depth", 0.0)
        ws["E16"] = data.get("spacing", 0.0)
        ws["E17"] = data.get("burden", 0.0)
        ws["E18"] = data.get("density", 0.0)

        # 🟢 Explosive Inputs (KG - you input final kg directly)
        ws["L9"] = float(data.get("watergel_per_hole", 0.0))
        ws["L10"] = float(data.get("watergel_per_sub_hole", 0.0))
        ws["L31"] = float(data.get("ammonium_per_hole", 0.0))
        ws["L32"] = float(data.get("ammonium_per_sub_hole", 0.0))

        # 🟢 ED Pattern (L15–L24)
        ed_pattern = data.get("ed_pattern", [0] * 10)
        for i in range(10):
            try:
                ws[f"L{15+i}"] = int(ed_pattern[i])
            except:
                ws[f"L{15+i}"] = ""

        # ✅ Save as .xlsx and return
        output_file = "filled_gi_form.xlsx"
        wb.save(output_file)

        return send_file(
            output_file,
            as_attachment=True,
            download_name="filled_gi_form.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        return {"error": f"PDF generation failed: {str(e)}"}, 500

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
