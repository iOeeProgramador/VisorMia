import pandas as pd
import zipfile
from datetime import datetime
from io import BytesIO
import streamlit as st

st.set_page_config(page_title="Combinador de Excel - Municipalidad", layout="wide")
st.title("📊 Generador de Datos Combinados (5 Archivos Excel)")

# ---- Subida de archivos ----
arch_ordenes = st.file_uploader("📂 Subir archivo ÓRDENES", type=["xlsx"], key="ordenes")
arch_stock = st.file_uploader("📂 Subir archivo STOCK", type=["xlsx"], key="stock")
arch_estado = st.file_uploader("📂 Subir archivo ESTADO", type=["xlsx"], key="estado")
arch_responsable = st.file_uploader("📂 Subir archivo RESPONSABLE", type=["xlsx"], key="resp")
arch_precios = st.file_uploader("📂 Subir archivo PRECIOS", type=["xlsx"], key="precio")

if all([arch_ordenes, arch_stock, arch_estado, arch_responsable, arch_precios]):

    try:
        # ---- Cargar datos ----
        df_ordenes = pd.read_excel(arch_ordenes)  # primera hoja
        xl_stock = pd.ExcelFile(arch_stock)
        df_bpcs = xl_stock.parse("BPCS")
        df_wms = xl_stock.parse("WMS")
        df_contenedor = xl_stock.parse("Contenedor")

        df_estado = pd.read_excel(arch_estado)
        df_responsable = pd.read_excel(arch_responsable, sheet_name="Empresa")
        df_precios = pd.read_excel(arch_precios)

        # ---- Mapas de referencia ----
        mapa_responsables = df_responsable.set_index("HNAME")["RESP"].to_dict()
        df_precios["LPROD"] = df_precios["LPROD"].astype(str).str.strip().str.upper()
        mapa_precios = df_precios.set_index("LPROD")["VALOR"].to_dict()

        # ---- Enriquecer datos de ordenes ----
        df_ordenes["RESP"] = df_ordenes["HNAME"].map(mapa_responsables)
        df_ordenes["LPROD"] = df_ordenes["LPROD"].astype(str).str.strip().str.upper()
        df_ordenes["VALOR"] = df_ordenes["LPROD"].map(mapa_precios)

        # ---- Crear archivo Excel combinado ----
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_ordenes.to_excel(writer, index=False, sheet_name="Ordenes")
            df_bpcs.to_excel(writer, index=False, sheet_name="BPCS")
            df_wms.to_excel(writer, index=False, sheet_name="WMS")
            df_contenedor.to_excel(writer, index=False, sheet_name="Contenedor")
            df_estado.to_excel(writer, index=False, sheet_name="Estado")
        output.seek(0)

        # ---- Vista previa interactiva ----
        st.success("✅ Archivos cargados correctamente. Vista previa de datos combinados:")

        st.write(f"🔢 Total de registros: {len(df_ordenes):,}")

        columnas_disponibles = df_ordenes.columns.tolist()

        columnas_seleccionadas = st.multiselect(
            "🧩 Selecciona las columnas que deseas visualizar:",
            options=columnas_disponibles,
            default=columnas_disponibles
        )

        st.dataframe(df_ordenes[columnas_seleccionadas], use_container_width=True)

        # ---- Botón de descarga Excel ----
        st.download_button(
            label="📥 Descargar archivo combinado (DatosCombinados.xlsx)",
            data=output,
            file_name="DatosCombinados.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.info("💡 Puedes exportar la tabla visible a PDF usando Ctrl+P desde el navegador.")

    except Exception as e:
        st.error(f"❌ Error al procesar los archivos: {e}")
