import pandas as pd
from datetime import datetime
from docxtpl import DocxTemplate
import os
import subprocess

doc = DocxTemplate("docs/carta-invitacion.docx")

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

hoy = f"{datetime.now().day} de {meses[datetime.now().month]} de {datetime.now().year}"
docente = "Mgst. Juan Carlos Cahuana quispe".title()
nombre_diplomado = "diplomado en educacion especial con mencion en trasntorno del espectro autista".upper()
nombre_modulo = "tecnologia del hormigon en instancias de construcción".upper()
competencia_modulo = "Analiza y aplica principios fundamentales de la tecnología del hormigón, incluyendo el diseño de mezclas y la evaluación de la durabilidad de las estructuras, para poder diseñar y especificar adecuadamente la composición del hormigón en función de las necesidades de cada proyecto, garantizando su resistencia, durabilidad y comportamiento frente a agentes externos."
contenidos_minimos = """
U.A.1. INTRODUCCIÓN AL MATERIAL HORMIGÓN Y SUS MATERIALES
1.1. Introducción al hormigón como material de construcción.
1.2. Componentes del hormigón: cemento, agua, áridos y aditivos.

U.A.2. CONSTITUYENTES, EVOLUCIÓN DEL HORMIGÓN, DESAFÍOS GLOBALES Y LOCALES, INNOVACIÓN EN HORMIGÓN
2.1. Evolución histórica de los constituyentes del hormigón.
2.2. Desafíos globales y locales en la construcción y su impacto en el desarrollo del hormigón.
2.3. Innovaciones en el diseño de mezclas y tecnologías para mejorar las propiedades del hormigón.

U.A.3. TECNOLOGÍA DEL HORMIGÓN
3.1. Procesos de mezcla, transporte, colocación y curado del hormigón fresco.
3.2. Técnicas de refuerzo y acabado del hormigón endurecido.

U.A.4. CONCEPTO DE DURABILIDAD. NORMAS INTERNACIONALES Y NCH170
4.1. Importancia de la durabilidad en las estructuras de hormigón.
4.2. Normas internacionales y NCh170 para el diseño y la construcción de estructuras durables de hormigón.

U.A.5. DISEÑO DE MEZCLAS
5.1. Principios y métodos para el diseño de mezclas de hormigón.
5.2. Consideraciones en la selección y proporción de materiales para obtener las propiedades deseadas del hormigón.

U.A.6. EJERCICIOS EN CLASES DE DISEÑO DE MEZCLAS
6.1. Aplicación práctica de los conocimientos adquiridos en el diseño de mezclas de hormigón.
6.2. Resolución de problemas prácticos y toma de decisiones informadas sobre la formulación y ajuste de mezclas de hormigón.
6.3. Ejercicios prácticos para mejorar la comprensión y habilidades en el diseño de mezclas.
"""

dias_clases = "Viernes y Sábado"
fechas_clases = [
    "24/11/2021",
    "25/11/2021",
    "01/12/2021",
    "02/12/2021",
    "08/12/2021",
    "09/12/2021",
]
objetivo = "Desarrollar competencias técnicas y conceptuales en el ámbito del hormigón armado, abordando aspectos relacionados con su tecnología, diseño, durabilidad, sustentabilidad, normativas aplicables, nuevas tecnologías de construcción y realización de monografías."
# data = {
#     "hoy": hoy,
#     "docente": docente,
#     "nombre_diplomado": nombre_diplomado,
#     "nombre_modulo": nombre_modulo,
#     "competencia_modulo": competencia_modulo,
#     "contenidos_minimos": contenidos_minimos,
#     "dias_clases": dias_clases,
#     "fechas_clases": fechas_clases,
#     "objetivo": objetivo,
# "fecha_clase_6": row["FechaClase6"].strftime("%d/%B/%Y"),
# }

df = pd.read_excel("GenerarCartasInvitacion.xlsx")

print(df)
for i, row in df.iterrows():
    data = {
        "hoy": hoy,
        "docente": row["Docente"].title(),
        "nombre_diplomado": row["DIPLOMADO"].upper(),
        "nombre_modulo": row["MODULO"].upper(),
        "competencia_modulo": row["COMPETENCIA DEL MODULO"],
        "contenidos_minimos": row["CONTENIDO MINIMO SUGERIDO"],
        "dias_clases": row["Dias de Clases"],
        "fecha_clase_1": row["FechaClase1"].strftime("%d/%m/%Y"),
        "fecha_clase_2": row["FechaClase2"].strftime("%d/%m/%Y"),
        "fecha_clase_3": row["FechaClase3"].strftime("%d/%m/%Y"),
        "fecha_clase_4": row["FechaClase4"].strftime("%d/%m/%Y"),
        "fecha_clase_5": row["FechaClase5"].strftime("%d/%m/%Y"),
        "fecha_clase_6": row["FechaClase6"].strftime("%d/%m/%Y"),
        "objetivo": row["OBJETIVO"],
    }
    # print(data)
    doc.render(data)
    docente = row["Docente"].replace(".", "").title()
    doc.save(f"docs/Carta Invitacion - {docente}.docx")

    subprocess.run(
        [
            "libreoffice",
            "--headless",
            "--convert-to",
            "pdf",
            f"docs/Carta Invitacion - {docente}.docx",
            "--outdir",
            "docs",
        ],
    )
    os.remove(f"docs/Carta Invitacion - {docente}.docx")
# doc.render(data)
# nombre_docente = docente.replace(".", "").replace(" ", "-").lower()
# doc.save(f"docs/carta-invitacion-{nombre_docente}.docx")
