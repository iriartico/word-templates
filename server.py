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


def download_file(doc, file_name):
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


@app.route("/generar-carta-invitacion", methods=["POST"])
def generar_carta_invitacion():
    data = request.json
    if not data:
        return jsonify({"error": "No se enviaron datos."}), 400

    doc = DocxTemplate("docs/carta-invitacion.docx")

    hoy = f"{datetime.now().day} de {meses[datetime.now().month]} de {datetime.now().year}"
    data["hoy"] = hoy
    for i in range(1, 7):
        data[f"fecha_clase_{i}"] = convert_date(data.get(f"fecha_clase_{i}"))

    doc.render(data)

    docente = data.get("docente", "Desconocido").replace(".", "").title()
    file_name = f"Carta Invitacion - {docente}.docx"

    return download_file(doc, file_name)


@app.route("/generar-certificado-docente", methods=["POST"])
def generar_certificado_docente():
    data = request.json
    # print(data)
    if not data:
        return jsonify({"error": "No se enviaron datos."}), 400

    doc = DocxTemplate("docs/certificado-docente.docx")

    hoy = f"{datetime.now().day} de {meses[datetime.now().month]} de {datetime.now().year}"
    data["hoy"] = hoy
    data["fecha_inicio"] = convert_date(data.get("fecha_inicio"))
    data["fecha_fin"] = convert_date(data.get("fecha_fin"))

    doc.render(data)

    docente = data.get("docente", "Desconocido").replace(".", "").title()
    file_name = f"Certificado Docente - {docente}.docx"

    return download_file(doc, file_name)


if __name__ == "__main__":
    print("Iniciando servidor Flask...")
    app.run(host="0.0.0.0", port=5000, debug=True)
