import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime
from io import BytesIO

st.set_page_config(page_title="Reporte de Pendientes", layout="wide")
st.title("📋 Reporte de Pendientes de Regularización Documentaria")

# --- Cargar datos desde archivos Excel locales ---
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
        fecha = datetime.strptime(fecha_str, "%d.%m.%Y").date()
        df = pd.read_excel(archivo, sheet_name="BASE TOTAL", dtype=str)
        df["ARCHIVO_ORIGEN"] = archivo
        df["FECHA_ARCHIVO"] = fecha
        dfs.append(df)
    return pd.concat(dfs, ignore_index=True)

df = cargar_datos()
df["FECHA_ARCHIVO"] = pd.to_datetime(df["FECHA_ARCHIVO"]).dt.date
df["STATUS A DETALLE"] = df["STATUS A DETALLE"].str.upper()

# --- Filtrar pendientes (todos los días) ---
df_pendientes_total = df[df["STATUS A DETALLE"] != "COMPLETADO"].copy()

# --- Filtrar solo los del último día para la tabla ---
fecha_max = df["FECHA_ARCHIVO"].max()
df_pendientes = df_pendientes_total[df_pendientes_total["FECHA_ARCHIVO"] == fecha_max].copy()

# --- Filtros para tabla ---
col1, col2, col3, col4, col5, col6 = st.columns(6)
with col1:
    region = st.selectbox("🌎 REGIÓN", ["Todas"] + sorted(df_pendientes["REGIÓN"].dropna().unique()))
with col2:
    subregion = st.selectbox("🗺️ SUB.REGIÓN", ["Todas"] + sorted(df_pendientes["SUB.REGIÓN"].dropna().unique()))
with col3:
    locacion = st.selectbox("🏢 LOCACIÓN", ["Todas"] + sorted(df_pendientes["LOCACIÓN"].dropna().unique()))
with col4:
    mesa = st.selectbox("MESA", ["Todas"] + sorted(df_pendientes["MESA"].dropna().unique()))
with col5:
    ruta = st.selectbox("🛣️ RUTA", ["Todas"] + sorted(df_pendientes["RUTA"].dropna().astype(str).unique()))
with col6:
    codigo = st.selectbox("🔢 CÓDIGO", ["Todos"] + sorted(df_pendientes["CÓDIGO"].dropna().unique()))

# --- Aplicar filtros tabla ---
if region != "Todas":
    df_pendientes = df_pendientes[df_pendientes["REGIÓN"] == region]
if subregion != "Todas":
    df_pendientes = df_pendientes[df_pendientes["SUB.REGIÓN"] == subregion]
if locacion != "Todas":
    df_pendientes = df_pendientes[df_pendientes["LOCACIÓN"] == locacion]
if mesa != "Todas":
    df_pendientes = df_pendientes[df_pendientes["MESA"] == mesa]
if ruta != "Todas":
    df_pendientes = df_pendientes[df_pendientes["RUTA"].astype(str) == ruta]
if codigo != "Todos":
    df_pendientes = df_pendientes[df_pendientes["CÓDIGO"] == codigo]

# --- Mostrar tabla ---
st.markdown(f"🔍 {df_pendientes.shape[0]} pendientes encontrados")
st.dataframe(df_pendientes, use_container_width=True)

# --- Filtros exclusivos para gráfico ---
st.subheader("📊 Filtros para evolución de pendientes")
colg1, colg2, colg3, colg4, colg5 = st.columns(5)
with colg1:
    region_g = st.selectbox("🌎 REGIÓN", ["Todas"] + sorted(df_pendientes_total["REGIÓN"].dropna().unique()))
with colg2:
    subregion_g = st.selectbox("🗺️ SUB.REGIÓN", ["Todas"] + sorted(df_pendientes_total["SUB.REGIÓN"].dropna().unique()))
with colg3:
    locacion_g = st.selectbox("🏢 LOCACIÓN", ["Todas"] + sorted(df_pendientes_total["LOCACIÓN"].dropna().unique()))
with colg4:
    mesa_g = st.selectbox("MESA", ["Todas"] + sorted(df_pendientes_total["MESA"].dropna().unique()))
with colg5:
    ruta_g = st.selectbox("🛣️ RUTA", ["Todas"] + sorted(df_pendientes_total["RUTA"].dropna().astype(str).unique()))

# --- Aplicar filtros al dataset total ---
df_grafico = df_pendientes_total.copy()
if region_g != "Todas":
    df_grafico = df_grafico[df_grafico["REGIÓN"] == region_g]
if subregion_g != "Todas":
    df_grafico = df_grafico[df_grafico["SUB.REGIÓN"] == subregion_g]
if locacion_g != "Todas":
    df_grafico = df_grafico[df_grafico["LOCACIÓN"] == locacion_g]
if mesa_g != "Todas":
    df_grafico = df_grafico[df_grafico["MESA"] == mesa_g]
if ruta_g != "Todas":
    df_grafico = df_grafico[df_grafico["RUTA"].astype(str) == ruta_g]

# --- Agrupar datos para el gráfico ---
df_chart = df_grafico.groupby("FECHA_ARCHIVO")["CÓDIGO"].count().reset_index(name="PENDIENTES")

# --- Mostrar gráfico ---
if not df_chart.empty:
    st.subheader("📈 Evolución de pendientes por fecha")
    fig = px.line(df_chart, x="FECHA_ARCHIVO", y="PENDIENTES", markers=True)
    fig.update_layout(xaxis_title="Fecha", yaxis_title="Cantidad de Pendientes")
    st.plotly_chart(fig, use_container_width=True)
else:
    st.info("No hay datos disponibles para mostrar el gráfico.")

# --- Exportar Excel con formato y gráfico ---
def exportar_excel_formateado(df_export, nombre_hoja, incluir_grafico=False):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_export.to_excel(writer, index=False, sheet_name=nombre_hoja)
        workbook = writer.book
        worksheet = writer.sheets[nombre_hoja]

        # Estilo de encabezado
        header_format = workbook.add_format({'bold': True, 'bg_color': '#FFFF99', 'border': 1})
        for col_num, value in enumerate(df_export.columns.values):
            worksheet.write(0, col_num, value, header_format)
            worksheet.set_column(col_num, col_num, 20)

        # Gráfico
        if incluir_grafico:
            chart_data = df_export.groupby("FECHA_ARCHIVO")["CÓDIGO"].count().reset_index(name="PENDIENTES")
            chart_sheet = workbook.add_worksheet("Gráfico")
            chart_sheet.write_column("A2", chart_data["FECHA_ARCHIVO"].astype(str))
            chart_sheet.write_column("B2", chart_data["PENDIENTES"])
            chart = workbook.add_chart({'type': 'line'})
            chart.add_series({
                'categories': ['Gráfico', 1, 0, len(chart_data), 0],
                'values':     ['Gráfico', 1, 1, len(chart_data), 1],
                'name':       'Pendientes',
            })
            chart.set_title({'name': 'Evolución de pendientes'})
            chart.set_x_axis({'name': 'Fecha'})
            chart.set_y_axis({'name': 'Cantidad'})
            chart.set_legend({'position': 'bottom'})
            chart_sheet.insert_chart("D2", chart)

    return output.getvalue()

# --- Descargar pendientes filtrados ---
excel_data1 = exportar_excel_formateado(df_pendientes, "Pendientes")
st.download_button(
    label="📥 Descargar Excel filtrado",
    data=excel_data1,
    file_name="pendientes_filtrados.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

# --- Descargar evolución con gráfico ---
excel_data2 = exportar_excel_formateado(df_grafico, "EvolucionPendientes", incluir_grafico=True)
st.download_button(
    label="📥 Descargar evolución de pendientes",
    data=excel_data2,
    file_name="evolucion_pendientes.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)