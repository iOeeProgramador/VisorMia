import streamlit as st
import pandas as pd
import zipfile
import io
from datetime import datetime

st.set_page_config(layout="wide")
st.title("Procesador de archivos Excel desde ZIP")

uploaded_file = st.file_uploader("Carga tu archivo ZIP con los libros de Excel", type="zip")

if uploaded_file is not None:
    with zipfile.ZipFile(uploaded_file) as z:
        expected_files = ["ORDENES.xlsx", "INVENTARIO.xlsx", "ESTADO.xlsx", "PRECIOS.xlsx", "GESTION.xlsx"]
        file_dict = {name: z.open(name) for name in expected_files if name in z.namelist()}

        if "ORDENES.xlsx" in file_dict:
            # Cargar ORDENES y renombrar columnas
            df_ordenes = pd.read_excel(file_dict["ORDENES.xlsx"])
            df_ordenes.columns = [f"{col}_ORDENES" for col in df_ordenes.columns]

            # Agregar columna CONTROL_DIAS
            if "LRDTE_ORDENES" in df_ordenes.columns:
                today = datetime.today()
                df_ordenes.insert(0, "CONTROL_DIAS", df_ordenes["LRDTE_ORDENES"].apply(lambda x: (datetime.strptime(str(int(x)), "%Y%m%d") - today).days))

            # Si existe INVENTARIO, combinar con ORDENES
            if "INVENTARIO.xlsx" in file_dict:
                df_inventario = pd.read_excel(file_dict["INVENTARIO.xlsx"])
                df_inventario.columns = [f"{col}_INVENTARIO" for col in df_inventario.columns]

                # Unir por LPROD_ORDENES y Cod. Producto_INVENTARIO
                df_combinado = pd.merge(
                    df_ordenes,
                    df_inventario,
                    left_on="LPROD_ORDENES",
                    right_on="Cod. Producto_INVENTARIO",
                    how="left"
                )
            else:
                df_combinado = df_ordenes

            # Guardar en Excel combinado
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_combinado.to_excel(writer, index=False, sheet_name='Datos')
            output.seek(0)

            # Mostrar datos
            st.subheader("Vista previa de DatosCombinados.xlsx")
            st.dataframe(df_combinado, use_container_width=True)

            # Botón para descarga
            st.download_button(
                label="Salir y descargar DatosCombinados.xlsx",
                data=output,
                file_name="DatosCombinados.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        else:
            st.error("El archivo ORDENES.xlsx no fue encontrado en el ZIP.")
