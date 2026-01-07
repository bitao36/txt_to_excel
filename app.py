import os
import re
import pandas as pd
from datetime import datetime
from collections import Counter
from flask import Flask, render_template, request, send_file, flash, redirect, url_for

app = Flask(__name__)
app.secret_key = "fichas_secret"

UPLOAD_FOLDER = "uploads"
OUTPUT_FOLDER = "outputs"

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

COLUMNAS = [
    "MFN",
    "SIGNATURA TOPOGRAFICA",
    "AUTOR PRINCIPAL",
    "TITULO/SUBTITULO",
    "MENCION RESPONSABILIDAD",
    "EDICION",
    "IMPRENTA",
    "DESCRIPCION FISICA"
]

def leer_archivo_txt(ruta):
    try:
        with open(ruta, "r", encoding="utf-8") as f:
            return f.read()
    except UnicodeDecodeError:
        with open(ruta, "r", encoding="latin-1") as f:
            return f.read()

def parsear_registros(texto):
    registros = []
    mfns = []

    bloques = re.split(r"\n(?=MFN:)", texto.strip())

    for bloque in bloques:
        registro = {campo: "" for campo in COLUMNAS}
        for linea in bloque.splitlines():
            linea = linea.strip()
            if not linea:
                continue
            if linea.startswith("MFN:"):
                mfn = linea.replace("MFN:", "").strip()
                registro["MFN"] = mfn
                mfns.append(mfn)
            else:
                partes = linea.split("\t", 1)
                if len(partes) == 2:
                    campo, valor = partes
                    if campo.strip() in registro:
                        registro[campo.strip()] = valor.strip()

        registros.append(registro)

    return registros, mfns

def obtener_rango_mfn(mfns):
    mfns_ordenados = sorted(mfns, key=lambda x: int(re.sub(r"\D", "", x)))
    return mfns_ordenados[0], mfns_ordenados[-1]

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        archivo = request.files.get("archivo")
        if not archivo or archivo.filename == "":
            flash("Debes seleccionar un archivo .txt")
            return redirect(url_for("index"))

        ruta_txt = os.path.join(UPLOAD_FOLDER, archivo.filename)
        archivo.save(ruta_txt)

        texto = leer_archivo_txt(ruta_txt)
        registros, mfns = parsear_registros(texto)

        if not registros:
            flash("No se encontraron registros válidos")
            return redirect(url_for("index"))

        mfn_inicio, mfn_fin = obtener_rango_mfn(mfns)
        timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        nombre_excel = f"lista_{timestamp}_MFN{mfn_inicio}-{mfn_fin}.xlsx"
        ruta_excel = os.path.join(OUTPUT_FOLDER, nombre_excel)

        df = pd.DataFrame(registros, columns=COLUMNAS)
        df.to_excel(ruta_excel, index=False)

        duplicados = [mfn for mfn, c in Counter(mfns).items() if c > 1]
        if duplicados:
            flash(f"MFN duplicados detectados: {', '.join(duplicados)}")

        return send_file(ruta_excel, as_attachment=True)

    return render_template("index.html")

if __name__ == "__main__":
    #app.run(debug=True)
    app.run(host="0.0.0.0", port=5000)
