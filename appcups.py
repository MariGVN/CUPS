from flask import Flask, render_template, request
import pandas as pd
import os

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER

if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)
def generar_analisis(resultados):
    patrones = {}
    for resultado in resultados:
        if not resultado["encontrado"]:
            continue
        for coincidencia in resultado["coincidencias"]:
            datos = coincidencia["datos"]
            for campo, valor in datos.items():
                if campo not in patrones:
                    patrones[campo] = {}
                if valor not in patrones[campo]:
                    patrones[campo][valor] = 0
                patrones[campo][valor] += 1
    resumen = []
    for campo, valores in patrones.items():
        for valor, cantidad in valores.items():
            if cantidad > 1:
                resumen.append({
                    "campo": campo,
                    "valor": valor,
                    "cantidad": cantidad
                })              
    return resumen

@app.route("/")
def inicio():

    archivos = []

    for archivo in os.listdir(UPLOAD_FOLDER):

        if archivo.endswith(".xlsx"):

            hojas = []

            try:
                with pd.ExcelFile(
                    os.path.join(
                        UPLOAD_FOLDER,
                        archivo
                    )
                ) as excel:
                    hojas = excel.sheet_names
            except Exception:
                pass

                hojas = excel.sheet_names

            except Exception:
                pass

            archivos.append({
                "nombre": archivo,
                "hojas": hojas
            })

    return render_template(
        "buscador.html",
        archivos=archivos
    )

@app.route("/subir", methods=["POST"])
def subir():

    archivo = request.files["excel"]

    if archivo.filename:

        ruta = os.path.join(
            UPLOAD_FOLDER,
            archivo.filename
        )

        archivo.save(ruta)

    return inicio()

@app.route("/buscar", methods=["POST"])
def buscar():

    cups_buscado = (
        request.form["cups"]
        .strip()
        .upper()
    )

    resultados = []

    for archivo in os.listdir(UPLOAD_FOLDER):

        if not archivo.endswith(".xlsx"):
            continue

        ruta = os.path.join(
            UPLOAD_FOLDER,
            archivo
        )

        try:

            hojas = pd.read_excel(
                ruta,
                sheet_name=None
            )

            for nombre_hoja, df in hojas.items():

                try:

                    # Caso especial datos generales
                    if nombre_hoja.upper() == "DATOS GENERALES":

                        df = pd.read_excel(
                            ruta,
                            sheet_name=nombre_hoja,
                            header=2
                        )
                    # Caso especial VF ELECTROMECANICOS ELIMINAR
                    if nombre_hoja.upper() == "VF ELECTROMECANICOS ELIMINAR":

                        df = pd.read_excel(
                            ruta,
                            sheet_name=nombre_hoja,
                            header=2
                        )

                    columna_cups = next(
                        (
                            col
                            for col in df.columns
                            if "CUPS" in str(col).upper()
                        ),
                        None
                    )

                    if columna_cups is None:
                        continue

                    df[columna_cups] = (
                        df[columna_cups]
                        .astype(str)
                        .str.strip()
                        .str.upper()
                    )

                    coincidencias = (
                        df[df[columna_cups] == cups_buscado]
                        
                    )

                    for _, fila in coincidencias.iterrows():

                        datos_fila = {}

                        for columna, valor in fila.items():

                            if pd.isna(valor):
                                continue

                            valor = str(valor).strip()

                            if valor == "":
                                continue

                            datos_fila[columna] = valor

                        resultados.append({
                            "archivo": archivo,
                            "hoja": nombre_hoja,
                            "datos": datos_fila
                        })

                except Exception as e:
                    print(
                        f"Error hoja {nombre_hoja}: {e}"
                    )

        except Exception as e:
            print(
                f"Error archivo {archivo}: {e}"
            )

    return render_template(
        "resultados.html",
        cups=cups_buscado,
        resultados=resultados
    )
@app.route("/eliminar/<nombre>")
def eliminar(nombre):

    ruta = os.path.join(
        UPLOAD_FOLDER,
        nombre
    )

    if os.path.exists(ruta):
        os.remove(ruta)

    return inicio()
@app.route("/eliminar_todos")
def eliminar_todos():

    for archivo in os.listdir(UPLOAD_FOLDER):

        if archivo.endswith(".xlsx"):

            ruta = os.path.join(
                UPLOAD_FOLDER,
                archivo
            )

            if os.path.exists(ruta):
                os.remove(ruta)

    return inicio()
@app.route("/analisis")
def pagina_analisis():
    return render_template(
        "analisis.html"
    )
@app.route("/analizar", methods=["POST"])
def analizar():

    cups_texto = request.form["cups"]

    lista_cups = [
        c.strip().upper()
        for c in cups_texto.splitlines()
        if c.strip()
    ]

    # Creacion de resultados una sola vez
    resultados_dict = {}

    for cups in lista_cups:
        resultados_dict[cups] = {
            "cups": cups,
            "encontrado": False,
            "coincidencias": []
        }

    # Recorre excel una vez
    for archivo in os.listdir(UPLOAD_FOLDER):

        if not archivo.endswith(".xlsx"):
            continue

        ruta = os.path.join(
            UPLOAD_FOLDER,
            archivo
        )

        try:

            hojas = pd.read_excel(
                ruta,
                sheet_name=None
            )

            for nombre_hoja, df in hojas.items():

                try:

                    # Hojas especiales (no empiezan como las demas hojas)
                    if nombre_hoja.upper() == "DATOS GENERALES":

                        df = pd.read_excel(
                            ruta,
                            sheet_name=nombre_hoja,
                            header=2
                        )

                    if nombre_hoja.upper() == "VF ELECTROMECANICOS ELIMINAR":

                        df = pd.read_excel(
                            ruta,
                            sheet_name=nombre_hoja,
                            header=2
                        )

                    columna_cups = next(
                        (
                            col
                            for col in df.columns
                            if "CUPS" in str(col).upper()
                        ),
                        None
                    )

                    if columna_cups is None:
                        continue

                    df[columna_cups] = (
                        df[columna_cups]
                        .astype(str)
                        .str.strip()
                        .str.upper()
                    )

                    # Buscar todos los CUPS en la hoja
                    for cups_buscado in lista_cups:

                        filas = df[
                            df[columna_cups] == cups_buscado
                        ]

                        if filas.empty:
                            continue

                        resultados_dict[cups_buscado]["encontrado"] = True

                        for _, fila in filas.iterrows():

                            datos = {}

                            for columna, valor in fila.items():

                                if pd.isna(valor):
                                    continue

                                valor = str(valor).strip()

                                if valor == "":
                                    continue

                                datos[columna] = valor

                            resultados_dict[cups_buscado]["coincidencias"].append({
                                "archivo": archivo,
                                "hoja": nombre_hoja,
                                "datos": datos
                            })

                except Exception as e:

                    print(
                        f"Error hoja {nombre_hoja}: {e}"
                    )

        except Exception as e:

            print(
                f"Error archivo {archivo}: {e}"
            )

    resultados_analisis = list(
        resultados_dict.values()
    )

    resumen = generar_analisis(
        resultados_analisis
    )

    return render_template(
        "analisis_resultados.html",
        resultados=resultados_analisis,
        analisis=resumen
    )
if __name__ == "__main__":
    app.run(debug=True)


