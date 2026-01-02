# app.py
# Streamlit: Limpieza + equivalencias (F/G) + Gramaje
# - Col J y K: MAYÚSCULAS sin acentos + corrige texto tipo "FÃ³rmula"
# - Col G (cantidad vendida) y F (título producto): crea Cantidad, UPC, Descripcion
# - Crea Gramaje (en gramos) desde el texto del producto (ej 1.5kg -> 1500)
# FIX CSV ParserError:
# - Detecta encoding (chardet)
# - Prueba separador elegido y autodetect (sep=None)
# - Usa engine="python" + on_bad_lines="skip" para filas rotas

import io
import re
import unicodedata
from typing import Optional

import pandas as pd
import streamlit as st
import chardet

st.set_page_config(page_title="Procesador Excel/CSV", layout="wide")

# ----------------------------
# Helpers de texto
# ----------------------------
def fix_mojibake(s: str) -> str:
    """Corrige mojibake típico: 'FÃ³rmula' -> 'Fórmula'."""
    if s is None:
        return ""
    s = str(s)
    if any(ch in s for ch in ["Ã", "Â", "�"]):
        try:
            s2 = s.encode("latin1", errors="ignore").decode("utf-8", errors="ignore")
            if s2 and (s2.count("�") <= s.count("�")):
                return s2
        except Exception:
            pass
    return s

def strip_accents_upper(s: str) -> str:
    s = fix_mojibake(s)
    s = s.strip()
    s = re.sub(r"\s+", " ", s)
    s = "".join(
        c for c in unicodedata.normalize("NFD", s)
        if unicodedata.category(c) != "Mn"
    )
    return s.upper()

def parse_gramaje_grams(text: str) -> Optional[int]:
    """Extrae gramaje en gramos del texto (última ocurrencia)."""
    if text is None:
        return None
    t = fix_mojibake(str(text)).lower()
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
    """Convierte letra Excel (A=0) a nombre real de columna."""
    idx = ord(letter.upper()) - ord("A")
    if 0 <= idx < len(df.columns):
        return df.columns[idx]
    return None

def to_int_safe(x) -> Optional[int]:
    if x is None:
        return None
    s = str(x).strip()
    if s == "":
        return None
    s = s.replace(",", "")
    try:
        return int(float(s))
    except Exception:
        return None

# ----------------------------
# Equivalencias (Producto -> pack, upc, desc)
# Llaves: MAYUS SIN ACENTOS + fix mojibake
# ----------------------------
RAW_EQUIVS = [
    ("Fórmula Crecelac Bebé 0-12 Meses 1500gr", 1, "7501468141043", "CRECELAC 0-12 M 1.5 KG"),
    ("LecheLak - Leche de Cabra en Polvo 340gr La Mejor Opción Para Toda la Familia Calidad y Frescura en Cada Porción", 1, "7501468144501", "LECHELAK LECHE DE CABRA 340 G"),
    ("6 Pack Fórmula Crecelac Bebé 0-12 Meses 800gr", 6, "7501468140442", "CRECELAC 0-12 M 800 GR"),
    ("6 Pack Fórmula Crecelac Firstep 1-3 Años 1500gr", 6, "7501468140947", "CRECELAC FIRSTEP 1-3 AÑOS 1.5 KG"),
    ("6 Pack Fórmula Crecelac Firstep 1-3 Años 800gr", 6, "7501468148301", "CRECELAC FIRSTEP 1-3 AÑOS 800 GR"),
    ("LecheLak - Leche de Cabra en Polvo 340gr La Mejor Opción Para Toda la Familia Calidad y Frescura en Cada Porción - 12 pack", 12, "7501468144501", "LECHELAK LECHE DE CABRA 340 G"),
    ("FÃ³rmula Crecelac Firstep 1-3 AÃ±os 360gr", 1, "7501468148103", "CRECELAC FIRSTEP 1-3 AÑOS 360 GR"),
    # Lo dejé EXACTO como lo pegaste (si quieres lo corregimos)
    ("FÃ³rmula Crecelac Firstep 1-3 AÃ±os 800gr", 1, "7501468140442", "CRECELAC 0-12 M 800 GR"),
    ("FÃ³rmula Crecelac BebÃ© 0-12 Meses 400gr", 1, "7501468145508", "CRECELAC 0-12 M 400 GR"),
    ("12 Pack Crecelac BebÃ© 0-12 Meses 400gr", 12, "7501468145508", "CRECELAC 0-12 M 400 GR"),
    ("Fórmula Crecelac Firstep 1-3 Años 1500gr", 1, "7501468140947", "CRECELAC FIRSTEP 1-3 AÑOS 1.5 KG"),
]

EQUIV = {}
for title, pack_qty, upc, desc in RAW_EQUIVS:
    EQUIV[strip_accents_upper(title)] = (int(pack_qty), str(upc), str(desc))

# ----------------------------
# UI
# ----------------------------
st.title("Procesador de Excel/CSV (J,K + equivalencias F/G + Gramaje)")

with st.expander("Qué hace este app", expanded=True):
    st.markdown(
        """
- **Columna J** → MAYÚSCULAS sin acentos (corrige `FÃ³rmula`).
- **Columna K** → igual que J.
- **Columna F (producto)** + **Columna G (unidades vendidas)**:
  - Busca equivalencia por el texto del producto.
  - Crea: **Cantidad**, **UPC**, **Descripcion**.
  - **Cantidad = G * pack (1/6/12)**.
- Crea **Gramaje** en gramos desde el texto: `1.5kg` → `1500`.
- Carga CSV robusta (autodetect + skip líneas rotas).
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
# Carga robusta
# ----------------------------
def load_csv_robusto(file, separador_csv: str) -> pd.DataFrame:
    raw = file.getvalue()
    detected = chardet.detect(raw)
    enc = detected.get("encoding") or "utf-8"

    attempts = [
        dict(sep=separador_csv, engine="python", encoding=enc),
        dict(sep=None, engine="python", encoding=enc),  # autodetect separador
        dict(sep=separador_csv, engine="python", encoding="latin1"),
        dict(sep=None, engine="python", encoding="latin1"),
    ]

    last_err = None
    for kw in attempts:
        try:
            return pd.read_csv(
                io.BytesIO(raw),
                dtype=str,
                keep_default_na=False,
                on_bad_lines="skip",
                **kw
            )
        except Exception as e:
            last_err = e

    raise last_err

def load_file(file) -> pd.DataFrame:
    name = file.name.lower()
    if name.endswith(".csv"):
        return load_csv_robusto(file, separador_csv)
    return pd.read_excel(file, dtype=str, keep_default_na=False)

df = load_file(uploaded)

st.subheader("Vista previa (antes)")
st.dataframe(df.head(20), use_container_width=True)

# ----------------------------
# Procesamiento
# ----------------------------
colJ = get_col_by_letter(df, "J")
colK = get_col_by_letter(df, "K")
colF = get_col_by_letter(df, "F")  # producto
colG = get_col_by_letter(df, "G")  # unidades vendidas

warnings = []
if colJ is None: warnings.append("No existe la columna **J** (por posición).")
if colK is None: warnings.append("No existe la columna **K** (por posición).")
if colF is None: warnings.append("No existe la columna **F** (producto, por posición).")
if colG is None: warnings.append("No existe la columna **G** (unidades, por posición).")

if warnings:
    st.warning("⚠️ " + " ".join(warnings))

# J/K: mayúsculas sin acentos
for c in [colJ, colK]:
    if c is not None:
        df[c] = df[c].apply(lambda x: strip_accents_upper(x) if str(x).strip() != "" else "")

# Nuevas columnas
for newc in ["Cantidad", "UPC", "Descripcion", "Gramaje", "Producto_normalizado", "Equivalencia_encontrada"]:
    if newc not in df.columns:
        df[newc] = ""

unmatched = []

if colF is not None and colG is not None:
    for i, row in df.iterrows():
        g_raw = row.get(colG, "")
        f_raw = row.get(colF, "")

        g = to_int_safe(g_raw)
        f_fixed = fix_mojibake(str(f_raw))
        f_key = strip_accents_upper(f_fixed)

        df.at[i, "Producto_normalizado"] = f_key

        grams = parse_gramaje_grams(f_fixed)
        df.at[i, "Gramaje"] = "" if grams is None else str(grams)

        if g is None or f_key == "":
            df.at[i, "Equivalencia_encontrada"] = ""
            continue

        eq = EQUIV.get(f_key)
        if eq is None:
            df.at[i, "Equivalencia_encontrada"] = "NO"
            unmatched.append((i, f_raw))
            continue

        pack_qty, upc, desc = eq
        cantidad_total = g * pack_qty

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
    st.error(f"Hay {len(unmatched)} filas donde **F** no hizo match con la tabla de equivalencias.")
    with st.expander("Ver productos no encontrados (primeros 200)"):
        st.write(pd.DataFrame(unmatched, columns=["fila_index", "F_original"]).head(200))
else:
    st.success("OK: las filas con F/G encontraron equivalencia (o venían vacías).")

# ----------------------------
# Descarga
# ----------------------------
def df_to_excel_bytes(dataframe: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    engines = ["openpyxl", "xlsxwriter"]
    last_err = None

    for eng in engines:
        try:
            output.seek(0)
            output.truncate(0)
            with pd.ExcelWriter(output, engine=eng) as writer:
                dataframe.to_excel(writer, index=False, sheet_name="Procesado")
            output.seek(0)
            return output.read()
        except Exception as e:
            last_err = e

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
        st.info("Cambia el formato de salida a CSV como alternativa.")
else:
    cbytes = df_to_csv_bytes(df)
    st.download_button(
        "⬇️ Descargar CSV procesado",
        data=cbytes,
        file_name="procesado.csv",
        mime="text/csv",
    )

st.caption("Tip: si tus columnas no están exactamente en F/G/J/K por posición, reordena el archivo o dime los nombres reales de columnas y lo adapto por encabezados.")
