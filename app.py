from flask import Flask, render_template, request
import pandas as pd

app = Flask(__name__)

archivo = "AYESA.xlsx"

def buscar_cups(cups_buscado):
    cups_buscado = cups_buscado.strip().upper()
    texto_resultado = ""
    hojas = pd.read_excel(
          archivo,
          sheet_name=None
          
     )
    texto_resultado = ""
    for nombre_hoja, df in hojas.items():
        try:
         
            if nombre_hoja == "DATOS GENERALES":
                df = pd.read_excel(
              archivo,
              sheet_name=nombre_hoja,
              header=2
         )
            columna_cups = next(
                (col for col in df.columns if "CUPS" in (col).upper()),
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
            resultado = df[df[columna_cups] == cups_buscado.strip().upper()
        ]
            resultado = resultado.drop_duplicates()


            if not resultado.empty:

                texto_resultado += f"\n{'='*60}\n"
                texto_resultado += f"HOJA: {nombre_hoja}\n"
                texto_resultado += resultado.to_string()
                texto_resultado += f"{'='*60}\n\n"
                
                for _, fila in resultado.iterrows():

                    for columna, valor in fila.items():

                        if pd.isna(valor):
                            continue
                        valor = str(valor).strip()

                        if valor == "":
                         continue
                    texto_resultado += f"{columna}: {valor}\n"

                texto_resultado += "\n"
        except Exception as e:
            texto_resultado += (
                f"\nError en hoja "
                f"{nombre_hoja}: {e}\n"
            )    
    if texto_resultado == "":
        texto_resultado = (
            f"No se encontro el CUPS"
            f"{cups_buscado}"
        )
    return texto_resultado
    
@app.route("/", methods=["GET", "POST"])
def inicio():
    resultado = ""
    if request.method == "POST":
        cups = request.form["cups"]
        resultado = buscar_cups(cups)
    return render_template(
        "index.html",
        resultado=resultado
    )
if __name__ == "__main__":
    app.run(debug=True)