# -*- mode: python ; coding: utf-8 -*-
import sys
import threading
import webbrowser
import streamlit as st
import streamlit.web.bootstrap
import polars as pl
import pandas as pd
from io import BytesIO

# --- Streamlit Configuration ---
import os
# Force server port and disable XSRF/CORS conflicts
os.environ["STREAMLIT_SERVER_PORT"] = "8501"
os.environ["STREAMLIT_SERVER_ENABLE_XSRF_PROTECTION"] = "false"
os.environ["STREAMLIT_SERVER_ENABLE_CORS"] = "false"

# Helper to open browser
def _open_browser():
    webbrowser.open_new(f"http://localhost:{os.environ['STREAMLIT_SERVER_PORT']}")

# If running as frozen executable, bootstrap Streamlit and open browser once
if getattr(sys, 'frozen', False):
    # Schedule browser to open after server starts
    threading.Timer(2.0, _open_browser).start()
    # Start Streamlit server with fixed port and disabled CORS/XSRF
    streamlit.web.bootstrap.run(
        'procesarfacturas.py',
        '',
        [
            '--server.port=8501',
            '--server.enableXsrfProtection=false',
            '--server.enableCORS=false'
        ],
        {}
    )
    sys.exit()

# --- Helper functions ---
def safe_float(val):
    try:
        if val is None or str(val).strip() == "":
            return 0.0
        val = str(val).replace(",", ".").replace(" ", "")
        return float(val)
    except Exception:
        return 0.0


# --- Classification Logic ---
def clasificar_fila(row: dict) -> str:
    residencia = str(row.get("Residencia", "")).strip().lower()
    concepto = str(row.get("Concepto", "")).strip().lower()
    suplidos = safe_float(row.get("Suplidos", 0))
    retencion = safe_float(row.get("Retención", 0))
    retencion_pct = safe_float(row.get("% Retención", 0))
    base2 = safe_float(row.get("Base 2", 0))
    base3 = safe_float(row.get("Base 3", 0))
    iva1 = safe_float(row.get("% IVA 1", 0))
    iva2 = safe_float(row.get("% IVA 2", 0))
    iva3 = safe_float(row.get("% IVA 3", 0))
    ivas = [iva1, iva2, iva3]

        # Nacional
    if "nacional" in residencia:
        if base2 != 0 or base3 != 0:
            return "2 Bases"
        
        if suplidos != 0:
            return "Suplidos"
        
        if retencion != 0 or retencion_pct != 0:
            return "Con Retención"
        
        if any(i in [21, 10] for i in ivas):
            return "IVA 21 & 10"
        
        if any(i == 0 for i in ivas):
            return "IVA 0%"
        
        elif any(i not in [0, 10, 21] and i != 0 for i in ivas):
            return "IVA EXTRAÑO"
        
        return "No Clasificadas"

    # UE + concepto SUBCONTRATAS
    if "ue" in residencia and "subcontrat" in concepto:
        return "SUBCONTRATAS"
    
    # UE que NO es subcontrata
    if "ue" in residencia and "subcontrat" not in concepto:
        if any(i == 0 for i in ivas):
            return "Extranjero y subcontratas 0% IVA"
        else:
            return "No Clasificadas"

    # Extranjero (todo lo que no es nacional ni UE)
    if "nacional" not in residencia and "ue" not in residencia:
        return "Extranjero y subcontratas 0% IVA"

    return "No Clasificadas"

def procesar_excel(file) -> BytesIO:
    try:
        df_pandas = pd.read_excel(file, header=2, dtype=str)
    except Exception as e:
        st.error(f"❌ Error al leer el archivo Excel: {e}")
        return None

    df_pandas.dropna(axis=1, how="all", inplace=True)
    if df_pandas.empty:
        st.warning("⚠️ El archivo está vacío o no contiene datos en la fila 3 en adelante.")
        return None

    df = pl.from_pandas(df_pandas)
    cols_num = [
        "Suplidos",
        "Retención",
        "% Retención",
        "Base 2",
        "Base 3",
        "% IVA 1",
        "% IVA 2",
        "% IVA 3"
    ]
    existing_cols = [c for c in cols_num if c in df.columns]
    df = df.with_columns([
        pl.col(c).cast(pl.Float64, strict=False) for c in existing_cols
    ])
    df = df.with_columns([
        pl.struct(df.columns).map_elements(clasificar_fila).alias("Pestaña")
    ])
    pestañas = df.select("Pestaña").unique().to_series().to_list()
    dfs_por_pestaña = {
        nombre: df.filter(pl.col("Pestaña") == nombre).drop("Pestaña")
        for nombre in pestañas
    }

    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for nombre, df_p in dfs_por_pestaña.items():
            df_p.to_pandas().to_excel(writer, sheet_name=nombre[:31], index=False)
    output.seek(0)
    return output

st.title("Clasificador de Facturas Para la Importación a A3")
st.write("Arrastra aquí tu fichero Excel (NO CAMBIES EL FORMATO DE NOVATRANS), y descárgalo ya clasificado para contabilidad:")

archivo = st.file_uploader("Arrastra aquí tu archivo .xlsx", type=["xlsx"])

if archivo:
    with st.spinner("Procesando..."):
        resultado = procesar_excel(archivo)
    if resultado:
        st.success("¡Hecho! Descarga el archivo clasificado aquí:")
        st.download_button(
            label="Descargar Excel clasificado",
            data=resultado,
            file_name="clasificado_para_importar.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    
    if st.button("¿No funciona?"):
        st.balloons()
