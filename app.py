# VisorMiaOk_corregido.py

import streamlit as st
import pandas as pd
import zipfile
import io
import datetime
import base64
from zipfile import ZipFile
from io import BytesIO

st.set_page_config(layout="wide")
st.markdown("## \U0001F4C1 VisorMia | " + datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"))

st.markdown("### \U0001F4E6 Carga automática desde archivo ZIP (5 Excel incluidos)")

zip_file = st.file_uploader("Sube el archivo .ZIP que contenga los 5 Excel", type="zip")

if zip_file:
    with zipfile.ZipFile(zip_file, "r") as archive:
        filenames = archive.namelist()

        required_files = ["Ordenes.xlsx", "Stock.xlsx", "Estado.xlsx", "Precios.xlsx", "Responsable.xlsx"]
        if not all(req in filenames for req in required_files):
            st.error("\u274C El ZIP no contiene todos los archivos requeridos.")
        else:
            def read_excel_from_zip(name, skip_rows=0):
                with archive.open(name) as file:
                    return pd.read_excel(file, skiprows=skip_rows)

            try:
                ordenes = read_excel_from_zip("Ordenes.xlsx")
                stock = read_excel_from_zip("Stock.xlsx", skip_rows=2)  # Eliminar filas 1 y 2
                estado = read_excel_from_zip("Estado.xlsx")
                precios = read_excel_from_zip("Precios.xlsx")
                responsable = read_excel_from_zip("Responsable.xlsx")

                # Fusionar todos
                df = ordenes.copy()
                df["Control-Dias"] = (pd.to_datetime(df["LRDTE"].astype(str), format="%Y%m%d") - pd.Timestamp.now().normalize()).dt.days
                df = df.merge(stock, left_on="LPROD", right_on="Cod. Producto", how="left")
                df = df.merge(estado, on=["LORD", "LLINE"], how="left")
                df = df.merge(precios, on="LPROD", how="left")
                df = df.merge(responsable, on="HNAME", how="left")

                # Guardar archivo combinado
                output = BytesIO()
                with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                    df.to_excel(writer, index=False, sheet_name="DatosCombinados")
                output.seek(0)

                st.download_button(
                    label="\U0001F4C4 Descargar archivo combinado (DatosCombinados.xlsx)",
                    data=output,
                    file_name="DatosCombinados.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

                resumen = df["Resp"].value_counts().reset_index()
                resumen.columns = ["Responsable", "Cantidad"]
                total = resumen["Cantidad"].sum()
                resumen["Porcentaje"] = (resumen["Cantidad"] / total * 100).round(2)
                st.dataframe(resumen)

                if st.button("Mostrar Datos Detallados"):
                    mostrar_cols = [
                        "Control-Dias", "HEDTE", "HROUT", "LORD", "LLINE", "LPROD", "LDESC", "HNAME",
                        "Cod. Producto", "Ubicación", "Contenedor", "Zona", "Sitio", "pedido",
                        "UNICO", "OBSERVACION", "Valor", "On Hand", "Resp"
                    ]
                    df_visible = df[mostrar_cols].copy()
                    st.dataframe(df_visible, use_container_width=True)
            except Exception as e:
                st.error(f"\u274C Error al procesar el archivo: {str(e)}")
