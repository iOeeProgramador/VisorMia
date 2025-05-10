import streamlit as st
import pandas as pd
import zipfile
import io
from datetime import datetime
import streamlit.components.v1 as components

st.set_page_config(layout="wide")
st.title("Procesador de archivos MIA")

# Control de navegación entre páginas
menu = st.sidebar.selectbox("Navegación", ["Resumen y Datos", "Gestión de Columnas"])

uploaded_file = st.file_uploader("Carga tu archivo ZIP con los libros de Excel", type="zip")

if uploaded_file is not None:
    with zipfile.ZipFile(uploaded_file) as z:
        expected_files = ["ORDENES.xlsx", "INVENTARIO.xlsx", "ESTADO.xlsx", "PRECIOS.xlsx", "GESTION.xlsx"]
        file_dict = {name: z.open(name) for name in expected_files if name in z.namelist()}

        if "ORDENES.xlsx" in file_dict:
            df_ordenes = pd.read_excel(file_dict["ORDENES.xlsx"])
            df_ordenes.columns = [f"{col}_ORDENES" for col in df_ordenes.columns]

            if "LRDTE_ORDENES" in df_ordenes.columns:
                today = datetime.today()
                df_ordenes.insert(0, "CONTROL_DIAS", df_ordenes["LRDTE_ORDENES"].apply(lambda x: (datetime.strptime(str(int(x)), "%Y%m%d") - today).days))

            if "INVENTARIO.xlsx" in file_dict:
                df_inventario = pd.read_excel(file_dict["INVENTARIO.xlsx"])
                df_inventario.columns = [f"{col}_INVENTARIO" for col in df_inventario.columns]
                df_inventario_unique = df_inventario.drop_duplicates(subset=["Cod. Producto_INVENTARIO"])
                df_combinado = pd.merge(df_ordenes, df_inventario_unique, left_on="LPROD_ORDENES", right_on="Cod. Producto_INVENTARIO", how="left")
            else:
                df_combinado = df_ordenes

            if "ESTADO.xlsx" in file_dict:
                df_estado = pd.read_excel(file_dict["ESTADO.xlsx"])
                df_estado.columns = [f"{col}_ESTADO" for col in df_estado.columns]
                df_combinado["KEY_ORDENES"] = df_combinado["LORD_ORDENES"].astype(str) + df_combinado["LLINE_ORDENES"].astype(str)
                df_estado["KEY_ESTADO"] = df_estado["LORD_ESTADO"].astype(str) + df_estado["LLINE_ESTADO"].astype(str)
                df_estado_unique = df_estado.drop_duplicates(subset=["KEY_ESTADO"])
                df_combinado = pd.merge(df_combinado, df_estado_unique, left_on="KEY_ORDENES", right_on="KEY_ESTADO", how="left")

            if "PRECIOS.xlsx" in file_dict:
                df_precios = pd.read_excel(file_dict["PRECIOS.xlsx"])
                df_precios.columns = [f"{col}_PRECIOS" for col in df_precios.columns]
                df_precios_unique = df_precios.drop_duplicates(subset=["LPROD_PRECIOS"])
                for col in ["VALOR_PRECIOS", "On Hand_PRECIOS"]:
                    if col in df_precios_unique.columns:
                        df_precios_unique[col] = pd.to_numeric(df_precios_unique[col], errors='coerce').fillna(0).astype(int)
                df_combinado = pd.merge(df_combinado, df_precios_unique, left_on="LPROD_ORDENES", right_on="LPROD_PRECIOS", how="left")

            if "GESTION.xlsx" in file_dict:
                df_gestion = pd.read_excel(file_dict["GESTION.xlsx"])
                df_gestion.columns = [f"{col}_GESTION" for col in df_gestion.columns]
                df_gestion_unique = df_gestion.drop_duplicates(subset=["HNAME_GESTION"])
                df_combinado = pd.merge(df_combinado, df_gestion_unique, left_on="HNAME_ORDENES", right_on="HNAME_GESTION", how="left")

            # Diccionario de colores por sufijo
            color_dict = {
                "ORDENES": "#e6f7ff",
                "INVENTARIO": "#e6ffe6",
                "ESTADO": "#fff5e6",
                "PRECIOS": "#ffe6f0",
                "GESTION": "#f2e6ff",
                "CONTROL_DIAS": "#f9f9f9"
            }

            if menu == "Resumen y Datos":
                if "RESPONSABLE_GESTION" in df_combinado.columns:
                    resumen = df_combinado.groupby("RESPONSABLE_GESTION", dropna=False).size().reset_index(name="Total Líneas")
                    resumen["RESPONSABLE_GESTION"] = resumen["RESPONSABLE_GESTION"].fillna("SIN RESPONSABLE")
                    resumen = resumen.sort_values(by="Total Líneas", ascending=False)
                    total = resumen["Total Líneas"].sum()
                    st.subheader(f"Resumen Total de Líneas por Responsable (Total: {total})")
                    st.dataframe(resumen, use_container_width=True)

                st.subheader("Vista previa de DatosCombinados.xlsx")
                st.dataframe(df_combinado, use_container_width=True)

                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df_combinado.to_excel(writer, index=False, sheet_name='Datos')
                output.seek(0)

                st.download_button(
                    label="Salir y descargar DatosCombinados.xlsx",
                    data=output,
                    file_name="DatosCombinados.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            elif menu == "Gestión de Columnas":
                st.subheader("Gestión de Columnas - Mostrar/Ocultar")
                todas_las_columnas = df_combinado.columns.tolist()
                columnas_visibles = st.multiselect("Selecciona las columnas a mostrar (conservarán su color)", options=todas_las_columnas, default=todas_las_columnas)

                def get_col_color(col):
                    for key in color_dict:
                        if col.endswith(f"_{key}") or col == key:
                            return color_dict[key]
                    return "#ffffff"

                styled_df = df_combinado.copy()
                styled_df = styled_df[columnas_visibles]

                def highlight_columns():
                    return pd.DataFrame(
                        [[f"background-color: {get_col_color(col)}" for col in styled_df.columns]],
                        columns=styled_df.columns
                    )

                st.dataframe(styled_df.style.apply(highlight_columns, axis=0), use_container_width=True)

        else:
            st.error("El archivo ORDENES.xlsx no fue encontrado en el ZIP.")
