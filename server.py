from flask import Flask, request, send_from_directory, jsonify
from docxtpl import DocxTemplate
from datetime import datetime, timedelta
import os
import subprocess
import tempfile


app = Flask(__name__)

meses = {
    1: "enero",
    2: "febrero",
    3: "marzo",
    4: "abril",
    5: "mayo",
    6: "junio",
    7: "julio",
    8: "agosto",
    9: "septiembre",
    10: "octubre",
    11: "noviembre",
    12: "diciembre",
}


def convert_date(value):
    # Google Sheets usa como referencia el 30 de diciembre de 1899
    if value:
        base_date = datetime(1899, 12, 30)
        converted_date = base_date + timedelta(days=int(value))
        return converted_date.strftime("%d/%m/%Y")
    else:
        return value


@app.route("/generar-carta-invitacion", methods=["POST"])
def generar_carta():
    data = request.json
    print(data)
    if not data:
        return jsonify({"error": "No se enviaron datos."}), 400

    doc = DocxTemplate("docs/carta-invitacion.docx")

    hoy = f"{datetime.now().day} de {meses[datetime.now().month]} de {datetime.now().year}"
    data["hoy"] = hoy
    data["fecha_clase_1"] = convert_date(data.get("fecha_clase_1"))
    data["fecha_clase_2"] = convert_date(data.get("fecha_clase_2"))
    data["fecha_clase_3"] = convert_date(data.get("fecha_clase_3"))
    data["fecha_clase_4"] = convert_date(data.get("fecha_clase_4"))
    data["fecha_clase_5"] = convert_date(data.get("fecha_clase_5"))
    data["fecha_clase_6"] = convert_date(data.get("fecha_clase_6"))

    doc.render(data)

    docente = data.get("docente", "Desconocido").replace(".", "").title()
    file_name = f"Carta Invitacion - {docente}.docx"

    with tempfile.TemporaryDirectory() as tmp_dir:
        file_path = os.path.join(tmp_dir, file_name)
        pdf_path = file_path.replace(".docx", ".pdf")
        doc.save(file_path)

        subprocess.run(
            [
                "libreoffice",
                "--headless",
                "--convert-to",
                "pdf",
                file_path,
                "--outdir",
                tmp_dir,
            ],
            check=True,
        )

        return send_from_directory(
            tmp_dir, os.path.basename(pdf_path), as_attachment=True
        )


if __name__ == "__main__":
    print("Iniciando servidor Flask...")
    app.run(host="0.0.0.0", port=5000, debug=True)
