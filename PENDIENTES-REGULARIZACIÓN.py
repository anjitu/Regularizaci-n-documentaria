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

# --- Filtrar todos los pendientes para evolución ---
df_pendientes_total = df[df["STATUS A DETALLE"] != "COMPLETADO"].copy()

# --- Filtrar solo los del último día para la tabla principal ---
fecha_max = df["FECHA_ARCHIVO"].max()
df_pendientes = df_pendientes_total[df_pendientes_total["FECHA_ARCHIVO"] == fecha_max].copy()

# --- Filtros para tabla (cascada) ---
col1, col2, col3, col4, col5, col6 = st.columns(6)

region = col1.selectbox("🌎 REGIÓN", ["Todas"] + sorted(df_pendientes["REGIÓN"].dropna().unique()), key="region_t")
df_temp = df_pendientes[df_pendientes["REGIÓN"] == region] if region != "Todas" else df_pendientes
subregion = col2.selectbox("🗺️ SUB.REGIÓN", ["Todas"] + sorted(df_temp["SUB.REGIÓN"].dropna().unique()), key="subregion_t")
df_temp = df_temp[df_temp["SUB.REGIÓN"] == subregion] if subregion != "Todas" else df_temp
locacion = col3.selectbox("🏢 LOCACIÓN", ["Todas"] + sorted(df_temp["LOCACIÓN"].dropna().unique()), key="locacion_t")
df_temp = df_temp[df_temp["LOCACIÓN"] == locacion] if locacion != "Todas" else df_temp
mesa = col4.selectbox("💼 MESA", ["Todas"] + sorted(df_temp["MESA"].dropna().unique()), key="mesa_t")
df_temp = df_temp[df_temp["MESA"] == mesa] if mesa != "Todas" else df_temp
ruta = col5.selectbox("🚣️ RUTA", ["Todas"] + sorted(df_temp["RUTA"].dropna().astype(str).unique()), key="ruta_t")
df_temp = df_temp[df_temp["RUTA"].astype(str) == ruta] if ruta != "Todas" else df_temp
codigo = col6.selectbox("🔢 CÓDIGO", ["Todos"] + sorted(df_temp["CÓDIGO"].dropna().unique()), key="codigo_t")
df_pendientes = df_temp[df_temp["CÓDIGO"] == codigo] if codigo != "Todos" else df_temp

# --- Mostrar tabla de pendientes filtrada ---
st.markdown(f"🔍 {df_pendientes.shape[0]} pendientes encontrados")
st.dataframe(df_pendientes, use_container_width=True)

# --- Filtros exclusivos para gráfico ---
st.subheader("📊 Filtros para evolución de pendientes")
colg1, colg2, colg3, colg4, colg5 = st.columns(5)

region_g = colg1.selectbox("🌎 REGIÓN (Gráfico)", ["Todas"] + sorted(df_pendientes_total["REGIÓN"].dropna().unique()), key="region_g")
df_temp_g = df_pendientes_total[df_pendientes_total["REGIÓN"] == region_g] if region_g != "Todas" else df_pendientes_total
subregion_g = colg2.selectbox("🗺️ SUB.REGIÓN (Gráfico)", ["Todas"] + sorted(df_temp_g["SUB.REGIÓN"].dropna().unique()), key="subregion_g")
df_temp_g = df_temp_g[df_temp_g["SUB.REGIÓN"] == subregion_g] if subregion_g != "Todas" else df_temp_g
locacion_g = colg3.selectbox("🏢 LOCACIÓN (Gráfico)", ["Todas"] + sorted(df_temp_g["LOCACIÓN"].dropna().unique()), key="locacion_g")
df_temp_g = df_temp_g[df_temp_g["LOCACIÓN"] == locacion_g] if locacion_g != "Todas" else df_temp_g
mesa_g = colg4.selectbox("💼 MESA (Gráfico)", ["Todas"] + sorted(df_temp_g["MESA"].dropna().unique()), key="mesa_g")
df_temp_g = df_temp_g[df_temp_g["MESA"] == mesa_g] if mesa_g != "Todas" else df_temp_g
ruta_g = colg5.selectbox("🚣️ RUTA (Gráfico)", ["Todas"] + sorted(df_temp_g["RUTA"].dropna().astype(str).unique()), key="ruta_g")
df_grafico = df_temp_g[df_temp_g["RUTA"].astype(str) == ruta_g] if ruta_g != "Todas" else df_temp_g

# --- Evolución de pendientes por fecha para gráfico ---
df_evolutivo = df_grafico.groupby("FECHA_ARCHIVO").size().reset_index(name="TOTAL_PENDIENTES")

if not df_evolutivo.empty:
    st.subheader("📈 Evolución de pendientes por fecha")
    fig = px.line(df_evolutivo, x="FECHA_ARCHIVO", y="TOTAL_PENDIENTES", markers=True)
    fig.update_layout(xaxis_title="Fecha", yaxis_title="Total de Pendientes")
    st.plotly_chart(fig, use_container_width=True)
else:
    st.info("No hay datos suficientes para mostrar el gráfico evolutivo.")