import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO
from datetime import datetime

st.set_page_config(page_title="Reporte de Pendientes", layout="wide")
st.title("📋 Reporte de Pendientes de Regularización Documentaria")

# --- Cargar datos desde archivos locales ---
@st.cache_data
def cargar_datos():
    archivos = [
        "CEO-LISTA DE PENDIENTES-09.06.2025.xlsx",
        "NORTE-LISTA DE PENDIENTES-09.06.2025.xlsx",
        "LIMA-LISTA DE PENDIENTES-09.06.2025.xlsx",
        "SUR-LISTA DE PENDIENTES-09.06.2025.xlsx",
        "CEO-LISTA DE PENDIENTES-10.06.2025.xlsx",
        "NORTE-LISTA DE PENDIENTES-10.06.2025.xlsx",
        "LIMA-LISTA DE PENDIENTES-10.06.2025.xlsx",
        "SUR-LISTA DE PENDIENTES-10.06.2025.xlsx"
    ]
    dfs = []
    for archivo in archivos:
        fecha_str = archivo.split("-")[-1].replace(".xlsx", "")
        try:
            fecha = datetime.strptime(fecha_str, "%d.%m.%Y").date()
        except:
            fecha = None
        df = pd.read_excel(archivo, sheet_name="BASE TOTAL", dtype=str)
        df["ARCHIVO_ORIGEN"] = archivo
        df["FECHA_ARCHIVO"] = fecha
        dfs.append(df)
    return pd.concat(dfs, ignore_index=True)

# --- Datos cargados ---
df = cargar_datos()
df["FECHA_ARCHIVO"] = pd.to_datetime(df["FECHA_ARCHIVO"]).dt.date

# --- Solo pendientes ---
df_pendientes = df[df["STATUS A DETALLE"].str.upper() != "COMPLETADO"].copy()

# --- Filtrar solo por última fecha ---
ultima_fecha = df_pendientes["FECHA_ARCHIVO"].max()
df_pendientes_ultima = df_pendientes[df_pendientes["FECHA_ARCHIVO"] == ultima_fecha].copy()

# --- Filtros ---
col1, col2, col3, col4, col5, col6 = st.columns(6)

with col1:
    region = st.selectbox("🌎 REGIÓN", ["Todas"] + sorted(df_pendientes_ultima["REGIÓN"].dropna().unique()))
if region != "Todas":
    df_pendientes_ultima = df_pendientes_ultima[df_pendientes_ultima["REGIÓN"] == region]

with col2:
    subregion = st.selectbox("🗺️ SUB.REGIÓN", ["Todas"] + sorted(df_pendientes_ultima["SUB.REGIÓN"].dropna().unique()))
if subregion != "Todas":
    df_pendientes_ultima = df_pendientes_ultima[df_pendientes_ultima["SUB.REGIÓN"] == subregion]

with col3:
    locacion = st.selectbox("🏢 LOCACIÓN", ["Todas"] + sorted(df_pendientes_ultima["LOCACIÓN"].dropna().unique()))
if locacion != "Todas":
    df_pendientes_ultima = df_pendientes_ultima[df_pendientes_ultima["LOCACIÓN"] == locacion]

with col4:
    mesa = st.selectbox("💼 MESA", ["Todas"] + sorted(df_pendientes_ultima["MESA"].dropna().unique()))
if mesa != "Todas":
    df_pendientes_ultima = df_pendientes_ultima[df_pendientes_ultima["MESA"] == mesa]

with col5:
    ruta = st.selectbox("🛣️ RUTA", ["Todas"] + sorted(df_pendientes_ultima["RUTA"].dropna().astype(str).unique()))
if ruta != "Todas":
    df_pendientes_ultima = df_pendientes_ultima[df_pendientes_ultima["RUTA"].astype(str) == ruta]

with col6:
    codigo = st.text_input("🔢 CÓDIGO", "")
if codigo:
    df_pendientes_ultima = df_pendientes_ultima[df_pendientes_ultima["CÓDIGO"].str.contains(codigo, na=False)]

# --- Tabla final ---
st.markdown(f"🔍 {len(df_pendientes_ultima)} pendientes encontrados")
st.dataframe(df_pendientes_ultima, use_container_width=True)

# --- Evolutivo para gráfico ---
df_filtrado_total = df[df["STATUS A DETALLE"].str.upper() != "COMPLETADO"].copy()
if region != "Todas":
    df_filtrado_total = df_filtrado_total[df_filtrado_total["REGIÓN"] == region]
if subregion != "Todas":
    df_filtrado_total = df_filtrado_total[df_filtrado_total["SUB.REGIÓN"] == subregion]
if locacion != "Todas":
    df_filtrado_total = df_filtrado_total[df_filtrado_total["LOCACIÓN"] == locacion]
if mesa != "Todas":
    df_filtrado_total = df_filtrado_total[df_filtrado_total["MESA"] == mesa]
if ruta != "Todas":
    df_filtrado_total = df_filtrado_total[df_filtrado_total["RUTA"].astype(str) == ruta]
if codigo:
    df_filtrado_total = df_filtrado_total[df_filtrado_total["CÓDIGO"].str.contains(codigo, na=False)]

# --- Agrupar evolución ---
df_evolutivo = df_filtrado_total.groupby([
    "FECHA_ARCHIVO", "REGIÓN", "SUB.REGIÓN", "LOCACIÓN", "MESA", "RUTA"
]).size().reset_index(name="TOTAL_PENDIENTES")

# --- Gráfico en app ---
if not df_evolutivo.empty:
    st.subheader("📈 Evolución de pendientes por fecha")
    df_chart = df_evolutivo.groupby("FECHA_ARCHIVO")["TOTAL_PENDIENTES"].sum().reset_index()
    fig = px.line(df_chart, x="FECHA_ARCHIVO", y="TOTAL_PENDIENTES", markers=True)
    fig.update_layout(xaxis_title="Fecha", yaxis_title="Total de Pendientes")
    st.plotly_chart(fig, use_container_width=True)
else:
    st.info("No hay datos suficientes para mostrar el gráfico evolutivo.")

# --- Función exportar Excel con gráfico y formato ---
def exportar_con_formato(df_export):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_export.to_excel(writer, index=False, sheet_name="EvolucionPendientes")
        workbook = writer.book
        worksheet = writer.sheets["EvolucionPendientes"]

        # Formato de encabezado
        header_format = workbook.add_format({'bold': True, 'bg_color': '#DCE6F1', 'border': 1})
        for col_num, value in enumerate(df_export.columns.values):
            worksheet.write(0, col_num, value, header_format)
            worksheet.set_column(col_num, col_num, len(value) + 5)

        # Insertar gráfico
        chart = workbook.add_chart({'type': 'line'})
        chart.add_series({
            'categories': ['EvolucionPendientes', 1, 0, len(df_export), 0],
            'values':     ['EvolucionPendientes', 1, df_export.columns.get_loc("TOTAL_PENDIENTES"), len(df_export), df_export.columns.get_loc("TOTAL_PENDIENTES")],
            'name':       'Pendientes por Fecha',
            'marker':     {'type': 'circle', 'size': 6}
        })
        chart.set_title({'name': '📊 Evolución de Pendientes'})
        chart.set_x_axis({'name': 'Fecha'})
        chart.set_y_axis({'name': 'Total Pendientes'})
        worksheet.insert_chart("I2", chart)
    return output.getvalue()

# --- Botón para descargar evolución con formato y gráfico ---
st.download_button(
    label="📥 Descargar evolución de pendientes (formateado)",
    data=exportar_con_formato(df_evolutivo),
    file_name="evolucion_pendientes.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

# --- Botón para descargar pendientes filtrados ---
def to_excel_bytes(df_export):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_export.to_excel(writer, index=False, sheet_name="Pendientes")
    return output.getvalue()

st.download_button(
    label="📥 Descargar pendientes filtrados",
    data=to_excel_bytes(df_pendientes_ultima),
    file_name="pendientes_filtrados.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)