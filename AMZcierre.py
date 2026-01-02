  # app.py
# Streamlit: Limpieza + equivalencias (Q/R) + Gramaje
# - Col J y K: MAYÚSCULAS sin acentos + corrige texto tipo "FÃ³rmula"
# - Col Q (cantidad vendida) y R (título producto): crea Cantidad, UPC, Descripcion
# - Crea Gramaje (en gramos) desde el texto del producto (ej 1.5kg -> 1500)

import io
import re
import unicodedata
from typing import Optional, Tuple

import pandas as pd
import streamlit as st

st.set_page_config(page_title="Procesador Excel/CSV", layout="wide")

# ----------------------------
# Helpers de texto
# ----------------------------
def fix_mojibake(s: str) -> str:
    """Corrige casos típicos de mojibake: 'FÃ³rmula' -> 'Fórmula'."""
    if s is None:
        return ""
    s = str(s)
    if any(ch in s for ch in ["Ã", "Â", "�"]):
        try:
            s2 = s.encode("latin1", errors="ignore").decode("utf-8", errors="ignore")
            # Si el resultado se ve mejor, úsalo
            if s2 and (s2.count("�") < s.count("�")):
                return s2
        except Exception:
            pass
    return s

def strip_accents_upper(s: str) -> str:
    s = fix_mojibake(s)
    s = s.strip()
    # Normaliza espacios
    s = re.sub(r"\s+", " ", s)
    # Quita acentos
    s = "".join(
        c for c in unicodedata.normalize("NFD", s)
        if unicodedata.category(c) != "Mn"
    )
    return s.upper()

def parse_gramaje_grams(text: str) -> Optional[int]:
    """
    Extrae el gramaje en gramos del texto.
    Regla: toma la ÚLTIMA ocurrencia que tenga unidad (g/gr/kg).
    Ej: "0-12 Meses 1500gr" -> 1500
        "1.5kg" -> 1500
    """
    if text is None:
        return None
    t = fix_mojibake(str(text)).lower()

    # Busca patrones con unidad
    # Ej: 1500gr, 800 g, 1.5kg, 340gr
    matches = list(re.finditer(r"(\d+(?:[.,]\d+)?)\s*(kg|kilo|kilogramos|g|gr|gramos)\b", t))
    if not matches:
        return None

    m = matches[-1]
    num = m.group(1).replace(",", ".")
    unit = m.group(2)

    try:
        val = float(num)
    except Exception:
        return None

    if unit in ("kg", "kilo", "kilogramos"):
        grams = int(round(val * 1000))
    else:
        grams = int(round(val))

    return grams if grams > 0 else None

def get_col_by_letter(df: pd.DataFrame, letter: str) -> Optional[str]:
    """Convierte letra Excel (A=0) a nombre de columna real."""
    idx = ord(letter.upper()) - ord("A")
    if 0 <= idx < len(df.columns):
        return df.columns[idx]
    return None

# ----------------------------
# Equivalencias (R -> pack, upc, desc)
# Normalizamos las llaves con MAYUS SIN ACENTOS.
# ----------------------------
RAW_EQUIVS = [
    ("Fórmula Crecelac Bebé 0-12 Meses 1500gr", 1, "7501468141043", "CRECELAC 0-12 M 1.5 KG"),
    ("LecheLak - Leche de Cabra en Polvo 340gr La Mejor Opción Para Toda la Familia Calidad y Frescura en Cada Porción", 1, "7501468144501", "LECHELAK LECHE DE CABRA 340 G"),
    ("6 Pack Fórmula Crecelac Bebé 0-12 Meses 800gr", 6, "7501468140442", "CRECELAC 0-12 M 800 GR"),
    ("6 Pack Fórmula Crecelac Firstep 1-3 Años 1500gr", 6, "7501468140947", "CRECELAC FIRSTEP 1-3 AÑOS 1.5 KG"),
    ("6 Pack Fórmula Crecelac Firstep 1-3 Años 800gr", 6, "7501468148301", "CRECELAC FIRSTEP 1-3 AÑOS 800 GR"),
    ("LecheLak - Leche de Cabra en Polvo 340gr La Mejor Opción Para Toda la Familia Calidad y Frescura en Cada Porción - 12 pack", 12, "7501468144501", "LECHELAK LECHE DE CABRA 340 G"),
    ("FÃ³rmula Crecelac Firstep 1-3 AÃ±os 360gr", 1, "7501468148103", "CRECELAC FIRSTEP 1-3 AÑOS 360 GR"),
    # Nota: tal como lo pegaste: mapeo a 7501468140442 / 800gr
    ("FÃ³rmula Crecelac Firstep 1-3 AÃ±os 800gr", 1, "7501468140442", "CRECELAC 0-12 M 800 GR"),
    ("FÃ³rmula Crecelac BebÃ© 0-12 Meses 400gr", 1, "7501468145508", "CRECELAC 0-12 M 400 GR"),
    ("12 Pack Crecelac BebÃ© 0-12 Meses 400gr", 12, "7501468145508", "CRECELAC 0-12 M 400 GR"),
    ("Fórmula Crecelac Firstep 1-3 Años 1500gr", 1, "7501468140947", "CRECELAC FIRSTEP 1-3 AÑOS 1.5 KG"),
]

EQUIV = {}
for title, pack_qty, upc, desc in RAW_EQUIVS:
    key = strip_accents_upper(title)
    EQUIV[key] = (int(pack_qty), str(upc), str(desc))

# ----------------------------
# UI
# ----------------------------
st.title("Procesador de Excel/CSV (J,K + equivalencias Q/R + Gramaje)")

with st.expander("Qué hace este app", expanded=True):
    st.markdown(
        """
- **Columna J** → convierte a **MAYÚSCULAS** y **sin acentos** (también corrige textos tipo `FÃ³rmula`).
- **Columna K** → igual que J.
- **Columna Q (unidades vendidas)** + **Columna R (título del producto)**:
  - Busca el producto en una tabla de equivalencias.
  - Crea 3 columnas nuevas: **Cantidad**, **UPC**, **Descripcion**.
  - **Cantidad = Q * (pack de 1/6/12)** según equivalencia.
- Crea **Gramaje** (en gramos) leyendo el gramaje del texto (ej. `1.5kg` → `1500`).
        """
    )

uploaded = st.file_uploader("Sube tu archivo (Excel .xlsx/.xls o CSV)", type=["xlsx", "xls", "csv"])

colA, colB = st.columns([1, 1])
with colA:
    salida_formato = st.selectbox("Formato de salida", ["Excel (.xlsx)", "CSV (.csv)"], index=0)
with colB:
    separador_csv = st.selectbox("Separador CSV (si subes CSV)", [",", ";", "|", "\t"], index=0)

if not uploaded:
    st.stop()

# ----------------------------
# Carga archivo
# ----------------------------
def load_file(file) -> pd.DataFrame:
    name = file.name.lower()
    if name.endswith(".csv"):
        return pd.read_csv(file, sep=separador_csv, dtype=str, keep_default_na=False)
    else:
        # Excel
        return pd.read_excel(file, dtype=str, keep_default_na=False)

df = load_file(uploaded)

st.subheader("Vista previa (antes)")
st.dataframe(df.head(20), use_container_width=True)

# ----------------------------
# Procesamiento
# ----------------------------
colJ = get_col_by_letter(df, "J")
colK = get_col_by_letter(df, "K")
colQ = get_col_by_letter(df, "Q")
colR = get_col_by_letter(df, "R")

warnings = []
if colJ is None:
    warnings.append("No existe la columna **J** (por posición).")
if colK is None:
    warnings.append("No existe la columna **K** (por posición).")
if colQ is None:
    warnings.append("No existe la columna **Q** (por posición).")
if colR is None:
    warnings.append("No existe la columna **R** (por posición).")

if warnings:
    st.warning("⚠️ " + " ".join(warnings))

# J/K: mayúsculas sin acentos
for c in [colJ, colK]:
    if c is not None:
        df[c] = df[c].apply(lambda x: strip_accents_upper(x) if str(x).strip() != "" else "")

# Nuevas columnas
df["Cantidad"] = ""
df["UPC"] = ""
df["Descripcion"] = ""
df["Gramaje"] = ""
df["Producto_normalizado"] = ""
df["Equivalencia_encontrada"] = ""

def to_int_safe(x) -> Optional[int]:
    if x is None:
        return None
    s = str(x).strip()
    if s == "":
        return None
    # elimina comas de miles
    s = s.replace(",", "")
    try:
        return int(float(s))
    except Exception:
        return None

unmatched = []

if colQ is not None and colR is not None:
    for i, row in df.iterrows():
        q_raw = row.get(colQ, "")
        r_raw = row.get(colR, "")

        q = to_int_safe(q_raw)
        r_fixed = fix_mojibake(str(r_raw))
        r_key = strip_accents_upper(r_fixed)

        df.at[i, "Producto_normalizado"] = r_key

        # Gramaje desde el texto original de R (ya corregido mojibake)
        grams = parse_gramaje_grams(r_fixed)
        df.at[i, "Gramaje"] = "" if grams is None else str(grams)

        if q is None or r_key == "":
            df.at[i, "Equivalencia_encontrada"] = ""
            continue

        eq = EQUIV.get(r_key)
        if eq is None:
            df.at[i, "Equivalencia_encontrada"] = "NO"
            unmatched.append((i, r_raw))
            continue

        pack_qty, upc, desc = eq
        cantidad_total = q * pack_qty

        df.at[i, "Cantidad"] = str(cantidad_total)
        df.at[i, "UPC"] = upc
        df.at[i, "Descripcion"] = desc
        df.at[i, "Equivalencia_encontrada"] = "SI"

# ----------------------------
# Resultados
# ----------------------------
st.subheader("Vista previa (después)")
st.dataframe(df.head(50), use_container_width=True)

if unmatched:
    st.error(f"Hay {len(unmatched)} filas donde **R** no hizo match con la tabla de equivalencias.")
    with st.expander("Ver productos no encontrados"):
        st.write(pd.DataFrame(unmatched, columns=["fila_index", "R_original"]).head(200))

# ----------------------------
# Descarga
# ----------------------------
def df_to_excel_bytes(dataframe: pd.DataFrame) -> bytes:
    output = io.BytesIO()

    # Intentar varios engines
    engines = ["openpyxl", "xlsxwriter"]
    last_err = None
    for eng in engines:
        try:
            with pd.ExcelWriter(output, engine=eng) as writer:
                dataframe.to_excel(writer, index=False, sheet_name="Procesado")
            output.seek(0)
            return output.read()
        except Exception as e:
            last_err = e
            output = io.BytesIO()

    # Si falla Excel, regresamos error claro
    raise RuntimeError(f"No se pudo generar Excel (.xlsx). Error: {last_err}")

def df_to_csv_bytes(dataframe: pd.DataFrame) -> bytes:
    return dataframe.to_csv(index=False).encode("utf-8-sig")

st.markdown("---")
st.subheader("Descargar")

if salida_formato.startswith("Excel"):
    try:
        xbytes = df_to_excel_bytes(df)
        st.download_button(
            "⬇️ Descargar Excel procesado",
            data=xbytes,
            file_name="procesado.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    except Exception as e:
        st.error(str(e))
        st.info("Como alternativa, cambia el formato de salida a CSV.")
else:
    cbytes = df_to_csv_bytes(df)
    st.download_button(
        "⬇️ Descargar CSV procesado",
        data=cbytes,
        file_name="procesado.csv",
        mime="text/csv",
    )

st.caption("Tip: si tus columnas no están exactamente en J/K/Q/R por posición, reordena el archivo o dime los nombres reales de las columnas y lo adapto a encabezados.")
