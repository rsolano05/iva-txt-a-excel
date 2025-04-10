import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog
from datetime import datetime

# === CONFIGURACIÓN DE ARCHIVO ===
PLANTILLA_EXCEL = "IVA TOTAL OFICIAL.xlsx"  # nombre de la plantilla oficial (debe estar en la misma carpeta)

# === VENTANA PARA SELECCIONAR ARCHIVO TXT ===
root = tk.Tk()
root.withdraw()
ruta_txt = filedialog.askopenfilename(title="Seleccione el archivo TXT", filetypes=[("Archivos de texto", "*.txt")])

if not ruta_txt:
    print("No se seleccionó ningún archivo.")
    exit()

# === LECTURA DEL TXT ===
columnas = [
    "Soc", "Ej", "Periodo", "Fecha.Doc", "Fe.contab.", "Cta", "Denominacion",
    "Ref", "Asignacion", "Doc.comp.", "Mon", "Importe base", "ctaimpto",
    "IVA reperc.pagar", "Importe bruto", "IVA repercutido"
]

df = pd.read_csv(
    ruta_txt,
    sep='|',
    header=None,
    names=columnas,
    dtype=str,
    encoding='utf-8',
    skip_blank_lines=True
)

# === LIMPIEZA Y TIPOS ===
def limpiar_numero(valor):
    if pd.isna(valor): return 0.0
    valor = str(valor).replace(".", "").replace(",", ".").replace(" ", "")
    try:
        return float(valor)
    except:
        return 0.0

# Columnas numéricas
col_numericas = ["Importe base", "IVA reperc.pagar", "Importe bruto", "IVA repercutido"]
for col in col_numericas:
    df[col] = df[col].apply(limpiar_numero)

# Columnas de texto que deben conservar ceros y no convertirse en notación científica
df["Ref"] = df["Ref"].astype(str).str.zfill(16)
df["Asignacion"] = df["Asignacion"].astype(str).str.zfill(16)

# Fecha contable al formato DD/MM/YYYY
df["Fe.contab."] = pd.to_datetime(df["Fe.contab."], format="%d.%m.%Y", errors='coerce')
df["Fe.contab."] = df["Fe.contab."].dt.strftime("%d/%m/%Y")

# === CARGAR PLANTILLA ===
plantilla_df = pd.read_excel(PLANTILLA_EXCEL, sheet_name=None)

# Crear copia del DataFrame principal
df_datos_convertidos = df.copy()

# === GENERAR HOJA "Resumen SAP" ===
resumen = df.groupby("Soc")[
    ["Importe base", "IVA reperc.pagar", "Importe bruto", "IVA repercutido"]
].sum().reset_index()

resumen["IVA Neto"] = resumen["IVA repercutido"] + resumen["IVA reperc.pagar"]

# === DETECTAR ERRORES (EJEMPLO SIMPLE) ===
errores = df[df["Ref"].str.strip() == ""].copy()
errores["Observación"] = "Campo Ref vacío"

# === GUARDAR RESULTADO EN EXCEL ===
nombre_salida = "IVA TOTAL.xlsx"
with pd.ExcelWriter(nombre_salida, engine="xlsxwriter") as writer:
    df_datos_convertidos.to_excel(writer, sheet_name="Datos Convertidos", index=False)
    resumen.to_excel(writer, sheet_name="Resumen SAP", index=False)
    errores.to_excel(writer, sheet_name="Informe", index=False)

print("Conversión finalizada")