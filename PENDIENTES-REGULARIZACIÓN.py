import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO
from datetime import datetime
import os

st.set_page_config(page_title="Reporte de Pendientes", layout="wide")
st.title("📋 Consulta de Pendientes de Regularización Documentaria")

# --- 1. Cargar todos los archivos con 'PENDIENTES' en el nombre ---
@st.cache_data

def cargar_archivos():
    archivos = [f for f in os.listdir() if "PENDIENTES" in f and f.endswith(".xlsx")]
    df_total = []
    for archivo in archivos:
        try:
            fecha_str = archivo.split("-")[-1].replace(".xlsx", "")
            fecha = datetime.strptime(fecha_str, "%d.%m.%Y").date()
        except:
            fecha = None
        df = pd.read_excel(archivo, sheet_name="BASE TOTAL", dtype=str)
        df["ARCHIVO_ORIGEN"] = archivo
        df["FECHA_ARCHIVO"] = fecha
        df_total.append(df)
    return pd.concat(df_total, ignore_index=True)

# --- 2. Cargar y limpiar los datos ---
df = cargar_archivos()
df["STATUS A DETALLE"] = df["STATUS A DETALLE"].str.upper()
df_pendientes = df[df["STATUS A DETALLE"] != "COMPLETADO"].copy()

# --- 3. Filtros en cascada ---
col1, col2, col3, col4, col5, col6 = st.columns(6)
with col1:
    region = st.selectbox("🌎 REGIÓN", [""] + sorted(df_pendientes["REGIÓN"].dropna().unique()))

filtro = df_pendientes[df_pendientes["REGIÓN"] == region] if region else df_pendientes
with col2:
    subregion = st.selectbox("🗺️ SUB.REGIÓN", [""] + sorted(filtro["SUB.REGIÓN"].dropna().unique()))

filtro = filtro[filtro["SUB.REGIÓN"] == subregion] if subregion else filtro
with col3:
    locacion = st.selectbox("🏢 LOCACIÓN", [""] + sorted(filtro["LOCACIÓN"].dropna().unique()))

filtro = filtro[filtro["LOCACIÓN"] == locacion] if locacion else filtro
with col4:
    mesa = st.selectbox("MESA", [""] + sorted(filtro["MESA"].dropna().unique()))

filtro = filtro[filtro["MESA"] == mesa] if mesa else filtro
with col5:
    ruta = st.selectbox("🛣️ RUTA", [""] + sorted(filtro["RUTA"].dropna().astype(str).unique()))

filtro = filtro[filtro["RUTA"].astype(str) == ruta] if ruta else filtro
with col6:
    codigo = st.selectbox("🧾 CÓDIGO", [""] + sorted(filtro["CÓDIGO"].dropna().astype(str).unique()))

filtro = filtro[filtro["CÓDIGO"].astype(str) == codigo] if codigo else filtro

# --- 4. Mostrar resultados ---
st.markdown(f"🔍 {len(filtro)} resultados encontrados")
st.dataframe(filtro, use_container_width=True)

# --- 5. Botón para exportar tabla filtrada ---
def to_excel_bytes(df_export):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_export.to_excel(writer, index=False, sheet_name='Pendientes')
        workbook = writer.book
        worksheet = writer.sheets['Pendientes']

        header_format = workbook.add_format({
            'bold': True, 'bg_color': '#4472C4', 'font_color': 'white', 'border': 1
        })

        for i, col in enumerate(df_export.columns):
            worksheet.set_column(i, i, max(len(col), 15))
            worksheet.write(0, i, col, header_format)
        worksheet.freeze_panes(1, 0)
        worksheet.autofilter(0, 0, df_export.shape[0], df_export.shape[1]-1)
    return output.getvalue()

excel_filtrado = to_excel_bytes(filtro)
st.download_button("📥 Descargar Excel Filtrado", data=excel_filtrado, file_name="pendientes_filtrados.xlsx",
                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# --- 6. Evolución de pendientes (agregada con región, sub, etc.) ---
campos = ["FECHA_ARCHIVO", "REGIÓN", "SUB.REGIÓN", "LOCACIÓN", "MESA", "RUTA"]
df_evolucion = df_pendientes.groupby(campos).size().reset_index(name="TOTAL_PENDIENTES")

# --- 7. Mostrar gráfico ---
if not filtro.empty:
    df_grafico = filtro.groupby("FECHA_ARCHIVO").size().reset_index(name="TOTAL_PENDIENTES")
    fig = px.line(df_grafico, x="FECHA_ARCHIVO", y="TOTAL_PENDIENTES", markers=True,
                  title="📈 Evolución de pendientes filtrados")
    fig.update_layout(xaxis_title="Fecha", yaxis_title="N° Pendientes", xaxis_tickformat='%d-%m-%Y')
    st.plotly_chart(fig, use_container_width=True)
else:
    st.info("No hay datos para mostrar gráfico con los filtros actuales.")

# --- 8. Botón para exportar evolución ---
excel_evolucion = to_excel_bytes(df_evolucion)
st.download_button("📥 Descargar Excel Evolutivo", data=excel_evolucion, file_name="evolucion_pendientes.xlsx",
                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
