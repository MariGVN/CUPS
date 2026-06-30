import pandas as pd
import tkinter as tk
from tkinter.scrolledtext import ScrolledText

archivo = "AYESA.xlsx"

cups_buscado = input("Introduce el CUPS: ").strip()

hojas = pd.read_excel(archivo, sheet_name=None)

texto_resultado = ""

for nombre_hoja, df in hojas.items():
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
    resultado = df[df[columna_cups] == cups_buscado
        ].drop_duplicates()


    if not resultado.empty:

            texto_resultado += f"\n{'='*60}\n"
            texto_resultado += f"HOJA: {nombre_hoja}\n"
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

if not texto_resultado:
    texto_resultado = f"No se encontró el CUPS {cups_buscado}"

# Crear ventana
root = tk.Tk()
root.title(f"Resultado CUPS {cups_buscado}")
root.geometry("900x700")

texto = ScrolledText(root, wrap=tk.WORD, font=("Consolas", 10))
texto.pack(fill="both", expand=True)

texto.insert("1.0", texto_resultado)
texto.config(state="disabled")

root.mainloop()

