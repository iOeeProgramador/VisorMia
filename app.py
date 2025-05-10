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

            # Si existe INVENTARIO, agregar columnas relacionadas sin duplicar filas
            if "INVENTARIO.xlsx" in file_dict:
                df_inventario = pd.read_excel(file_dict["INVENTARIO.xlsx"])
                df_inventario.columns = [f"{col}_INVENTARIO" for col in df_inventario.columns]

                # Eliminar duplicados en INVENTARIO por Cod. Producto
                df_inventario_unique = df_inventario.drop_duplicates(subset=["Cod. Producto_INVENTARIO"])

                # Hacer merge manteniendo estructura de ORDENES
                df_combinado = pd.merge(
                    df_ordenes,
                    df_inventario_unique,
                    left_on="LPROD_ORDENES",
                    right_on="Cod. Producto_INVENTARIO",
                    how="left"
                )
            else:
                df_combinado = df_ordenes

            # Si existe ESTADO, agregar columnas relacionadas sin duplicar filas
            if "ESTADO.xlsx" in file_dict:
                df_estado = pd.read_excel(file_dict["ESTADO.xlsx"])
                df_estado.columns = [f"{col}_ESTADO" for col in df_estado.columns]

                # Crear columnas de clave de combinación
                df_combinado["KEY_ORDENES"] = df_combinado["LORD_ORDENES"].astype(str) + df_combinado["LLINE_ORDENES"].astype(str)
                df_estado["KEY_ESTADO"] = df_estado["LORD_ESTADO"].astype(str) + df_estado["LLINE_ESTADO"].astype(str)

                # Eliminar duplicados en ESTADO por clave
                df_estado_unique = df_estado.drop_duplicates(subset=["KEY_ESTADO"])

                # Unir manteniendo las filas de ORDENES
                df_combinado = pd.merge(
                    df_combinado,
                    df_estado_unique,
                    left_on="KEY_ORDENES",
                    right_on="KEY_ESTADO",
                    how="left"
                )

            # Si existe PRECIOS, agregar columnas relacionadas sin duplicar filas
            if "PRECIOS.xlsx" in file_dict:
                df_precios = pd.read_excel(file_dict["PRECIOS.xlsx"])
                df_precios.columns = [f"{col}_PRECIOS" for col in df_precios.columns]

                # Eliminar duplicados en PRECIOS por LPROD
                df_precios_unique = df_precios.drop_duplicates(subset=["LPROD_PRECIOS"])

                # Convertir columnas VALOR y On Hand a enteros sin decimales si existen
                for col in ["VALOR_PRECIOS", "On Hand_PRECIOS"]:
                    if col in df_precios_unique.columns:
                        df_precios_unique[col] = pd.to_numeric(df_precios_unique[col], errors='coerce').fillna(0).astype(int)

                # Hacer merge manteniendo estructura de ORDENES
                df_combinado = pd.merge(
                    df_combinado,
                    df_precios_unique,
                    left_on="LPROD_ORDENES",
                    right_on="LPROD_PRECIOS",
                    how="left"
                )

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
