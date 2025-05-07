import pandas as pd
from datetime import datetime
from io import BytesIO
import zipfile
import streamlit as st

st.set_page_config(page_title="Combinador desde ZIP - Municipalidad", layout="wide")
st.title("📦 Carga automática desde archivo ZIP (5 Excel incluidos)")

# ---- Subida de archivo ZIP ----
archivo_zip = st.file_uploader("📂 Sube el archivo .ZIP que contenga los 5 Excel", type=["zip"])

def buscar_archivo(nombres, clave):
    for name in nombres:
        if clave in name.lower() and name.lower().endswith(".xlsx"):
            return name
    return None

if archivo_zip:
    try:
        with zipfile.ZipFile(archivo_zip) as zip_ref:
            lista_archivos = zip_ref.namelist()

            # Buscar archivos por nombre parcial
            archivo_ordenes = buscar_archivo(lista_archivos, "orden")
            archivo_stock = buscar_archivo(lista_archivos, "stock")
            archivo_estado = buscar_archivo(lista_archivos, "estado")
            archivo_responsable = buscar_archivo(lista_archivos, "respons")
            archivo_precios = buscar_archivo(lista_archivos, "precio")

            # Validar existencia
            if None in [archivo_ordenes, archivo_stock, archivo_estado, archivo_responsable, archivo_precios]:
                raise ValueError("Faltan uno o más archivos requeridos en el ZIP. Verifica los nombres.")

            # Leer archivos desde el ZIP
            df_ordenes = pd.read_excel(zip_ref.open(archivo_ordenes))
            xl_stock = pd.ExcelFile(zip_ref.open(archivo_stock))
            hojas = xl_stock.sheet_names
            hoja_stock = next((h for h in hojas if h.lower().startswith("stock")), None)
            hoja_wms = next((h for h in hojas if "wms" in h.lower()), None)
            hoja_contenedor = "Contenedor pendiente"
            if not hoja_stock or not hoja_wms or hoja_contenedor not in hojas:
                raise ValueError("No se encontraron todas las hojas requeridas en el archivo STOCK.")
            df_bpcs = xl_stock.parse(hoja_stock)
            df_wms = xl_stock.parse(hoja_wms)
            df_contenedor = xl_stock.parse(hoja_contenedor)

            df_estado = pd.read_excel(zip_ref.open(archivo_estado))
            df_responsable = pd.read_excel(zip_ref.open(archivo_responsable), sheet_name="Empresa")
            try:
                df_precios = pd.read_excel(zip_ref.open(archivo_precios))
            except Exception as e:
                raise ValueError("❌ No se pudo leer el archivo PRECIOS.xlsx. Verifica que esté bien formado.") from e

            # Procesamiento
            mapa_responsables = df_responsable.set_index("HNAME")["RESP"].to_dict()
            df_precios["LPROD"] = df_precios["LPROD"].astype(str).str.strip().str.upper()
            mapa_precios = df_precios.set_index("LPROD")["VALOR"].to_dict()
            df_ordenes["RESP"] = df_ordenes["HNAME"].map(mapa_responsables)
            df_ordenes["LPROD"] = df_ordenes["LPROD"].astype(str).str.strip().str.upper()
            df_ordenes["VALOR"] = df_ordenes["LPROD"].map(mapa_precios)

            # ---- Resumen por responsable con TOTAL ----
            resumen = df_ordenes["RESP"].value_counts().reset_index()
            resumen.columns = ["RESPONSABLE", "Total líneas"]
            resumen["Porcentaje"] = (resumen["Total líneas"] / len(df_ordenes) * 100).round(2)
            fila_total = pd.DataFrame({
                "RESPONSABLE": ["TOTAL"],
                "Total líneas": [resumen["Total líneas"].sum()],
                "Porcentaje": [100.0]
            })
            resumen = pd.concat([resumen, fila_total], ignore_index=True)
            resumen["Porcentaje"] = resumen["Porcentaje"].astype(str) + " %"

            st.subheader("📈 Resumen por Responsable")
            st.dataframe(resumen, use_container_width=True)

            # ---- Crear archivo combinado general ----
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_ordenes.to_excel(writer, index=False, sheet_name="Ordenes")
                df_bpcs.to_excel(writer, index=False, sheet_name="Stock")
                df_wms.to_excel(writer, index=False, sheet_name="WMS")
                df_contenedor.to_excel(writer, index=False, sheet_name="Contenedor pendiente")
                df_estado.to_excel(writer, index=False, sheet_name="Estado")
            output.seek(0)

            # ---- Vista previa interactiva ----
            st.success("✅ ZIP cargado correctamente. Vista previa de datos combinados:")
            st.write(f"🔢 Total de registros: {len(df_ordenes):,}")
            columnas_disponibles = df_ordenes.columns.tolist()
            columnas_seleccionadas = st.multiselect(
                "🧩 Selecciona las columnas que deseas visualizar:",
                options=columnas_disponibles,
                default=columnas_disponibles
            )
            st.dataframe(df_ordenes[columnas_seleccionadas], use_container_width=True)

            st.download_button(
                label="📥 Descargar archivo combinado (DatosCombinados.xlsx)",
                data=output,
                file_name="DatosCombinados.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            # ---- Archivos por responsable (ZIP) ----
            zip_buffer = BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                for responsable, grupo in df_ordenes.groupby("RESP"):
                    if pd.isna(responsable):
                        continue
                    file_buffer = BytesIO()
                    with pd.ExcelWriter(file_buffer, engine='openpyxl') as writer:
                        grupo.to_excel(writer, index=False, sheet_name="Ordenes")
                    file_buffer.seek(0)
                    nombre_archivo = f"RESP_{responsable.strip().replace(' ', '_')}.xlsx"
                    zip_file.writestr(nombre_archivo, file_buffer.read())
            zip_buffer.seek(0)

            st.download_button(
                label="📦 Descargar archivos por Responsable (ZIP)",
                data=zip_buffer,
                file_name="Archivos_por_Responsable.zip",
                mime="application/zip"
            )

            st.info("💡 Puedes exportar la tabla visible a PDF usando Ctrl+P desde el navegador.")

    except Exception as e:
        st.error(f"❌ Error al procesar el ZIP: {e}")
