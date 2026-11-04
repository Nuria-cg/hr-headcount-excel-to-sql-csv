#Python
Python script 

from pathlib import Path
import re
import sys
import unicodedata

import numpy as np
import pandas as pd


BASE_DIR = Path(__file__).resolve().parent


def resolver_ruta_entrada():
    if len(sys.argv) < 2:
        raise SystemExit("Uso: python LIMPIEZA_DATOS.py <archivo_excel>")

    ruta = Path(sys.argv[1])
    if not ruta.is_absolute():
        ruta = BASE_DIR / ruta

    if not ruta.exists():
        raise FileNotFoundError(f"No existe el archivo de entrada: {ruta}")

    return ruta


def construir_ruta_salida(ruta_entrada):
    output_dir = ruta_entrada.parent / "SQL"
    output_dir.mkdir(parents=True, exist_ok=True)
    return output_dir / f"{ruta_entrada.stem}.csv"


def quitar_acentos(valor):
    if not isinstance(valor, str):
        return valor

    resultado = []
    for caracter in unicodedata.normalize("NFD", valor):
        # NFD descompone ñ en n + tilde (U+0303): se conserva solo tras n/N
        if caracter == "̃" and resultado and resultado[-1] in "nN":
            resultado.append(caracter)
        elif unicodedata.category(caracter) != "Mn":
            resultado.append(caracter)

    return unicodedata.normalize("NFC", "".join(resultado))


def limpiar_cabecera(valor):
    if pd.isna(valor):
        return ""

    texto = str(valor).strip()
    texto = quitar_acentos(texto)
    texto = re.sub(r"[^0-9A-Za-zñÑ]+", "_", texto)
    texto = re.sub(r"_+", "_", texto).strip("_")
    return texto


def detectar_fila_cabecera(ruta_excel):
    preview = pd.read_excel(ruta_excel, engine="openpyxl", header=None, nrows=10)

    for indice, fila in preview.iterrows():
        valores = [limpiar_cabecera(valor) for valor in fila.tolist()]
        if "Empresa" in valores and "Nombre_empleado" in valores:
            return indice

    raise ValueError("No se ha encontrado la fila de cabeceras en el Excel")


# === Cargar Excel ============================================================
INPUT_FILE = resolver_ruta_entrada()
OUTPUT_FILE = construir_ruta_salida(INPUT_FILE)

header_row = detectar_fila_cabecera(INPUT_FILE)
df = pd.read_excel(INPUT_FILE, engine="openpyxl", header=header_row)

# 0️⃣ Limpiar cabeceras
df.columns = [limpiar_cabecera(col) for col in df.columns]

columnas_requeridas = [
    "Empresa",
    "Cotizacion_seguridad_soci",
    "Causa_de_la_baja",
    "Fecha_baja",
    "Fecha_alta",
    "Fecha_nacimiento",
    "Fecha_antiguedad",
]

faltantes = [col for col in columnas_requeridas if col not in df.columns]
if faltantes:
    raise ValueError(
        f"Faltan columnas requeridas: {faltantes}. "
        f"Columnas detectadas: {df.columns.tolist()}"
    )

# 1️⃣ Eliminar empresas no válidas
df = df[~df["Empresa"].astype(str).str.strip().isin(["1001", "Totales"])]

# 2️⃣ Normalizar códigos de empresa
df["Empresa"] = df["Empresa"].astype(str).str.strip().replace(
    {"1912": "24", "2281": "24"}
)

# 3️⃣ Eliminar empleados que no cotizan
df = df[
    df["Cotizacion_seguridad_soci"].astype(str).str.strip() != "No cotiza S.S."
]

# 4️⃣ Baja por fusión → limpiar datos de baja
mask_fusion = (
    df["Causa_de_la_baja"].astype(str).str.strip()
    == "Baja por fusión absorción empresa"
)
df.loc[mask_fusion, ["Fecha_baja", "Causa_de_la_baja"]] = np.nan

# 5️⃣ Fecha de baja sin causa → eliminar fecha
mask_fecha_baja = pd.to_datetime(df["Fecha_baja"], errors="coerce").notna()
mask_causa_vacia = df["Causa_de_la_baja"].isna() | (
    df["Causa_de_la_baja"].astype(str).str.strip() == ""
)
df.loc[mask_fecha_baja & mask_causa_vacia, "Fecha_baja"] = np.nan

# 6️⃣ Normalizar fechas
for col in [
    "Fecha_alta",
    "Fecha_nacimiento",
    "Fecha_antiguedad",
    "Fecha_baja",
]:
    df[col] = pd.to_datetime(df[col], errors="coerce").dt.strftime("%Y/%m/%d")

# 7️⃣ Eliminar acentos (conservando ñ)
df = df.map(quitar_acentos)

# 8️⃣ Convertir columnas numéricas a enteros
columnas_enteras = [
    "Empleado_Codigo",
    "Dias_alta_en_la_empresa_p",
    "Tarifa",
    "Codigo_categoria",
    "Codigo_contrato",
    "Dias_percepcion_I_T",
    "Dias_accidente",
]

for col in columnas_enteras:
    if col in df.columns:
        serie = pd.to_numeric(df[col], errors="coerce")
        df[col] = serie.map(
            lambda valor: "" if pd.isna(valor) else str(int(valor))
        )

# === Exportar CSV ============================================================
df.to_csv(OUTPUT_FILE, index=False, encoding="utf-8-sig")

print(f"OK: archivo generado en {OUTPUT_FILE}")

print(f"OK: archivo generado en {OUTPUT_FILE}")
