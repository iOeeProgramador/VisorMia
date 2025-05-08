import streamlit as st
import pandas as pd
import zipfile
import io
import datetime
from datetime import datetime as dt
from zipfile import BadZipFile

st.set_page_config(layout="wide")
st.title("📁 VisorMia | " + dt.now().strftime("%Y-%m-%d %H:%M:%S"))

st.subheader("📦 Carga automática desde archivo ZIP (5 Excel incluidos)")
st.caption("Sube el archivo .ZIP que contenga los 5 Excel")

archivo_zip = st.file_uploader("Drag and drop file here", type="zip")

COLUMNAS_REQUERIDAS = {
    "Ordenes": ["LPROD"],
    "Stock": ["Cod. Producto"],
    "Estado": ["LORD", "LLINE"],
    "Precios": ["LPROD"],
    "Responsable": ["HNAME"]
}

ESTADOS_EXTRA = [
    "B.O.", "CONFIRMACION EN 0", "CONTENEDOR R", "ENVIADO A RUTA 99", "FACTURACION PARCIAL",
    "FACTURADA", "FACTURADO", "GESTIONAR", "INBOUND", "PACKING", "PICKING", "POR FACTURAR",
    "PP EXTRAVIADO", "QUIEBRE", "REPLANIFICADO", "WMS"
]

archivos_excel = {}

def normalizar_columnas(df):
    df.columns = df.columns.str.strip().str.upper().str.replace(" ", "")
    return df

def validar_columnas(df, nombre_archivo, columnas_necesarias):
    faltantes = [col for col in columnas_necesarias if col.upper().replace(" ", "") not in df.columns]
    if faltantes:
        raise ValueError(f"Columnas faltantes en {nombre_archivo}: {', '.join(faltantes)}")

def mostrar_resumen_por_responsable(df):
    if "RESP" in df.columns:
        resumen = df["RESP"].value_counts().reset_index()
        resumen.columns = ["Responsable", "Total líneas"]
        resumen["Porcentaje"] = (resumen["Total líneas"] / len(df) * 100).round(2).astype(str) + " %"

        # Agregar columnas para los estados adicionales
        for estado in ESTADOS_EXTRA:
            col_estado = df[df["ESTADO"] == estado].groupby("RESP").size()
            resumen[estado] = resumen["Responsable"].map(col_estado).fillna(0).astype(int)

        st.subheader("📊 Resumen por Responsable")
        st.dataframe(resumen)

if archivo_zip:
    try:
        with zipfile.ZipFile(archivo_zip) as z:
            nombres_archivos = z.namelist()

            for nombre_logico, requeridas in COLUMNAS_REQUERIDAS.items():
                archivo_match = next((n for n in nombres_archivos if nombre_logico.lower() in n.lower()), None)
                if not archivo_match:
                    raise FileNotFoundError(f"No se encontró un archivo para: {nombre_logico}")

                with z.open(archivo_match) as f:
                    if nombre_logico == "Stock":
                        df = pd.read_excel(f, skiprows=2)
                    else:
                        df = pd.read_excel(f)

                    df = normalizar_columnas(df)
                    validar_columnas(df, nombre_logico, requeridas)
                    archivos_excel[nombre_logico] = df

            st.success("✅ Archivos cargados y validados correctamente.")

            # Mostrar resumen por responsable
            df_ordenes = archivos_excel.get("Ordenes")
            df_responsables = archivos_excel.get("Responsable")

            if df_ordenes is not None and df_responsables is not None:
                mapa_resp = df_responsables.set_index("HNAME")["RESP"].to_dict()
                df_ordenes["RESP"] = df_ordenes["HNAME"].map(mapa_resp)
                mostrar_resumen_por_responsable(df_ordenes)

    except BadZipFile:
        st.error("❌ El archivo subido no es un ZIP válido.")
    except FileNotFoundError as fe:
        st.error(f"❌ Error al procesar el ZIP: {str(fe)}")
    except ValueError as ve:
        st.error(f"❌ Error al procesar columnas: {str(ve)}")
    except Exception as e:
        st.error(f"❌ Error inesperado: {str(e)}")
