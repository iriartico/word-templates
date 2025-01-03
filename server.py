from flask import Flask, request, send_from_directory, jsonify
from docxtpl import DocxTemplate
from datetime import datetime
import os
import subprocess
import threading

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


@app.route("/generar-carta", methods=["POST"])
def generar_carta():
    data = request.json
    print(data)
    if not data:
        return jsonify({"error": "No se enviaron datos."}), 400

    # Plantilla base
    doc = DocxTemplate("docs/carta-invitacion.docx")

    # Fecha actual
    hoy = f"{datetime.now().day} de {meses[datetime.now().month]} de {datetime.now().year}"
    data["hoy"] = hoy

    # Renderizar el documento con los datos
    doc.render(data)

    # Nombre del archivo basado en docente
    docente = data.get("docente", "Desconocido").replace(".", "").title()
    file_name = f"Carta Invitacion - {docente}.docx"
    file_path = os.path.join("docs", file_name)
    pdf_path = file_path.replace(".docx", ".pdf")

    # Guardar el archivo DOCX
    doc.save(file_path)

    # Convertir a PDF con LibreOffice
    subprocess.run(
        [
            "libreoffice",
            "--headless",
            "--convert-to",
            "pdf",
            file_path,
            "--outdir",
            "docs",
        ]
    )

    # Eliminar el archivo DOCX para dejar solo el PDF
    os.remove(file_path)

    # Devolver el PDF generado
    return send_from_directory("docs", os.path.basename(pdf_path), as_attachment=True)


if __name__ == "__main__":
    # start_cloudflared(port=5000)

    print("Iniciando servidor Flask...")
    app.run(host="0.0.0.0", port=5000, debug=True)
