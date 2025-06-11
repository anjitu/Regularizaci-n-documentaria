import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime
from io import BytesIO

st.set_page_config(page_title="Reporte de Pendientes", layout="wide")
st.title("📋 Reporte de Pendientes de Regularización Documentaria")

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

# --- Cargar datos ---
df = cargar_datos()
df["FECHA_ARCHIVO"] = pd.to_datetime(df["FECHA_ARCHIVO"]).dt.date
df["STATUS A DETALLE"] = df["STATUS A DETALLE"].str.upper()

# --- Filtrar todos los pendientes ---
df_pendientes_total = df[df["STATUS A DETALLE"] != "COMPLETADO"].copy()

# --- Fecha más reciente para tabla ---
fecha_max = df["FECHA_ARCHIVO"].max()
df_pendientes = df_pendientes_total[df_pendientes_total["FECHA_ARCHIVO"] == fecha_max].copy()

# --- Filtros en cascada ---
col1, col2, col3, col4, col5, col6 = st.columns(6)

with col1:
    region = st.selectbox("🌎 REGIÓN", ["Todas"] + sorted(df["REGIÓN"].dropna().unique()), key="region")

with col2:
    subregion = st.selectbox("🗺️ SUB.REGIÓN", ["Todas"] + sorted(df[df["REGIÓN"] == region]["SUB.REGIÓN"].dropna().unique()) if region != "Todas" else ["Todas"] + sorted(df["SUB.REGIÓN"].dropna().unique()), key="subregion")

with col3:
    locacion = st.selectbox("🏢 LOCACIÓN", ["Todas"] + sorted(df[df["SUB.REGIÓN"] == subregion]["LOCACIÓN"].dropna().unique()) if subregion != "Todas" else ["Todas"] + sorted(df["LOCACIÓN"].dropna().unique()), key="locacion")

with col4:
    mesa = st.selectbox("💼 MESA", ["Todas"] + sorted(df[df["LOCACIÓN"] == locacion]["MESA"].dropna().unique()) if locacion != "Todas" else ["Todas"] + sorted(df["MESA"].dropna().unique()), key="mesa")

with col5:
    ruta = st.selectbox("🛣️ RUTA", ["Todas"] + sorted(df[df["MESA"] == mesa]["RUTA"].dropna().unique()) if mesa != "Todas" else ["Todas"] + sorted(df["RUTA"].dropna().unique()), key="ruta")

with col6:
    codigo = st.selectbox("🔢 CÓDIGO", ["Todos"] + sorted(df[df["RUTA"] == ruta]["CÓDIGO"].dropna().unique()) if ruta != "Todas" else ["Todos"] + sorted(df["CÓDIGO"].dropna().unique()), key="codigo")

# --- Aplicar filtros a ambas bases ---
filtros = {}
if region != "Todas": filtros["REGIÓN"] = region
if subregion != "Todas": filtros["SUB.REGIÓN"] = subregion
if locacion != "Todas": filtros["LOCACIÓN"] = locacion
if mesa != "Todas": filtros["MESA"] = mesa
if ruta != "Todas": filtros["RUTA"] = ruta
if codigo != "Todos": filtros["CÓDIGO"] = codigo

for col, val in filtros.items():
    df_pendientes = df_pendientes[df_pendientes[col] == val]
    df_pendientes_total = df_pendientes_total[df_pendientes_total[col] == val]

# --- Mostrar tabla ---
st.markdown(f"🔍 {df_pendientes.shape[0]} pendientes encontrados para la fecha {fecha_max}")
st.dataframe(df_pendientes, use_container_width=True)

# --- Generar base para gráfico evolutivo ---
df_evolutivo = (
    df_pendientes_total
    .groupby(["FECHA_ARCHIVO", "REGIÓN", "SUB.REGIÓN", "LOCACIÓN", "MESA", "RUTA"])
    .size()
    .reset_index(name="TOTAL_PENDIENTES")
)

df_chart = df_evolutivo.groupby("FECHA_ARCHIVO")["TOTAL_PENDIENTES"].sum().reset_index()

# --- Mostrar gráfico ---
if not df_chart.empty:
    st.subheader("📈 Evolución de pendientes por fecha")
    fig = px.line(df_chart, x="FECHA_ARCHIVO", y="TOTAL_PENDIENTES", markers=True)
    fig.update_layout(xaxis_title="Fecha", yaxis_title="Total de Pendientes")
    st.plotly_chart(fig, use_container_width=True)
else:
    st.info("No hay datos suficientes para mostrar el gráfico evolutivo.")

# --- Descargar Excel filtrado ---
def exportar_excel(df_export, nombre_hoja):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_export.to_excel(writer, index=False, sheet_name=nombre_hoja)
    return output.getvalue()

# Botón para descargar la tabla filtrada
excel_data1 = exportar_excel(df_pendientes, "Pendientes")
st.download_button(
    label="📥 Descargar Excel filtrado",
    data=excel_data1,
    file_name="pendientes_filtrados.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

# Botón para descargar la base del gráfico evolutivo
excel_data2 = exportar_excel(df_evolutivo, "EvolucionPendientes")
st.download_button(
    label="📥 Descargar evolución de pendientes",
    data=excel_data2,
    file_name="evolucion_pendientes.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)