pip install openpyxl
pip install pandas
pip install xlswriter
import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Reporte de Pendientes", layout="wide")
st.title("Consulta de Pendientes de Regularización Documentaria")

# --- 1. Carga automática de archivos ---
@st.cache_data
def cargar_datos():
    archivos = [
        "CEO-LISTA DE PENDIENTES-05.06.2025.xlsx",
        "NORTE-LISTA DE PENDIENTES-05.06.2025.xlsx",
        "LIMA-LISTA DE PENDIENTES-05.06.2025.xlsx",
        "SUR-LISTA DE PENDIENTES-05.06.2025.xlsx"
    ]
    dfs = []
    for archivo in archivos:
        df = pd.read_excel(archivo, sheet_name="BASE TOTAL", dtype=str)
        df["ARCHIVO_ORIGEN"] = archivo
        dfs.append(df)
    return pd.concat(dfs, ignore_index=True)

df = cargar_datos()

# --- 2. Filtrar solo pendientes ---
df_pendientes = df[df["STATUS A DETALLE"].str.upper() != "COMPLETADO"].copy()

# --- 3. Filtros dependientes ---
col1, col2, col3, col4, col5, col6 = st.columns(6)

with col1:
    region_opciones = [""] + sorted(df_pendientes["REGIÓN"].dropna().unique())
    region = st.selectbox("🌎 REGIÓN", region_opciones)

df_subreg = df_pendientes[df_pendientes["REGIÓN"] == region] if region else df_pendientes
with col2:
    subregion_opciones = [""] + sorted(df_subreg["SUB.REGIÓN"].dropna().unique())
    subregion = st.selectbox("🗺️ SUB.REGIÓN", subregion_opciones)

df_loc = df_subreg[df_subreg["SUB.REGIÓN"] == subregion] if subregion else df_subreg
with col3:
    locacion_opciones = [""] + sorted(df_loc["LOCACIÓN"].dropna().unique())
    locacion = st.selectbox("🏢 LOCACIÓN", locacion_opciones)

df_mesa = df_loc[df_loc["LOCACIÓN"] == locacion] if locacion else df_loc
with col4:
    mesa_opciones = [""] + sorted(df_mesa["MESA"].dropna().unique())
    mesa = st.selectbox("🍽️ MESA", mesa_opciones)

df_ruta = df_mesa[df_mesa["MESA"] == mesa] if mesa else df_mesa
with col5:
    ruta_opciones = [""] + sorted(df_ruta["RUTA"].dropna().astype(str).unique())
    ruta = st.selectbox("🛣️ RUTA", ruta_opciones)

with col6:
    codigo_cliente_opciones = [""] + sorted(df_ruta["CÓDIGO"].dropna().astype(str).unique())
    codigo_cliente = st.selectbox("🧾 CÓDIGO", codigo_cliente_opciones)

# --- 4. Aplicar filtros ---
df_filtrado = df_ruta.copy()
if ruta:
    df_filtrado = df_filtrado[df_filtrado["RUTA"].astype(str) == ruta]
if codigo_cliente:
    df_filtrado = df_filtrado[df_filtrado["CÓDIGO"].astype(str) == codigo_cliente]

# --- 5. Mostrar resultados ---
st.markdown(f"🔍 {len(df_filtrado)} resultados encontrados")
st.dataframe(df_filtrado, use_container_width=True)

# --- 6. Función para exportar Excel bonito ---
def to_excel_bytes(df_export):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_export.to_excel(writer, index=False, sheet_name='Pendientes')
        workbook  = writer.book
        worksheet = writer.sheets['Pendientes']

        # Formatos
        header_format = workbook.add_format({
            'bold': True,
            'bg_color': '#C00000',   # Rojo oscuro
            'font_color': 'white',
            'border': 1
        })

        yellow_header_format = workbook.add_format({
            'bold': True,
            'bg_color': '#FFF2CC',   # Amarillo claro
            'border': 1
        })

        # Aplicar formato a encabezado
        for col_num, column_name in enumerate(df_export.columns):
            if column_name.upper() == "STATUS A DETALLE":
                worksheet.write(0, col_num, column_name, yellow_header_format)
            else:
                worksheet.write(0, col_num, column_name, header_format)

        # Ajustar ancho de columnas
        for i, col in enumerate(df_export.columns):
            col_width = max(df_export[col].astype(str).map(len).max(), len(col)) + 2
            worksheet.set_column(i, i, col_width)

        # Congelar primera fila y activar filtros
        worksheet.freeze_panes(1, 0)
        worksheet.autofilter(0, 0, df_export.shape[0], df_export.shape[1] - 1)

    return output.getvalue()

# --- 7. Botón para descargar Excel ---
excel_data = to_excel_bytes(df_filtrado)
st.download_button(
    label="📥 Descargar Excel filtrado",
    data=excel_data,
    file_name="pendientes_filtrados.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
