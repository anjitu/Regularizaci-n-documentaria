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

# --- Crear base de pendientes no completados ---
df_pendientes_total = df[df["STATUS A DETALLE"] != "COMPLETADO"].copy()

# --- Mostrar solo los pendientes del último día para la tabla principal ---
fecha_max = df["FECHA_ARCHIVO"].max()
df_ultima_fecha = df_pendientes_total[df_pendientes_total["FECHA_ARCHIVO"] == fecha_max].copy()

# --- Filtros en cascada para la tabla ---
region = st.selectbox("🌎 REGIÓN", ["Todas"] + sorted(df["REGIÓN"].dropna().unique()), key="region")
if region != "Todas":
    df_ultima_fecha = df_ultima_fecha[df_ultima_fecha["REGIÓN"] == region]

subregion_options = ["Todas"] + sorted(df[df["REGIÓN"] == region]["SUB.REGIÓN"].dropna().unique()) if region != "Todas" else ["Todas"] + sorted(df["SUB.REGIÓN"].dropna().unique())
subregion = st.selectbox("🌏 SUB.REGIÓN", subregion_options, key="subregion")
if subregion != "Todas":
    df_ultima_fecha = df_ultima_fecha[df_ultima_fecha["SUB.REGIÓN"] == subregion]

locaciones = df["LOCACIÓN"].dropna().unique()
locacion = st.selectbox("🏢 LOCACIÓN", ["Todas"] + sorted(locaciones), key="locacion")
if locacion != "Todas":
    df_ultima_fecha = df_ultima_fecha[df_ultima_fecha["LOCACIÓN"] == locacion]

mesas = df["MESA"].dropna().unique()
mesa = st.selectbox("💼 MESA", ["Todas"] + sorted(mesas), key="mesa")
if mesa != "Todas":
    df_ultima_fecha = df_ultima_fecha[df_ultima_fecha["MESA"] == mesa]

rutas = df["RUTA"].dropna().astype(str).unique()
ruta = st.selectbox("🛣️ RUTA", ["Todas"] + sorted(rutas), key="ruta")
if ruta != "Todas":
    df_ultima_fecha = df_ultima_fecha[df_ultima_fecha["RUTA"].astype(str) == ruta]

# --- Filtro Código solo para tabla ---
codigos = df_ultima_fecha["CÓDIGO"].dropna().unique()
codigo = st.selectbox("🔢 CÓDIGO", ["Todos"] + sorted(codigos), key="codigo")
if codigo != "Todos":
    df_ultima_fecha = df_ultima_fecha[df_ultima_fecha["CÓDIGO"] == codigo]

# --- Mostrar tabla de pendientes ---
st.markdown(f"🔍 {df_ultima_fecha.shape[0]} pendientes encontrados (fecha {fecha_max})")
st.dataframe(df_ultima_fecha, use_container_width=True)

# --- Filtros para el gráfico de evolución ---
st.subheader("📈 Gráfico de Evolución de Pendientes")
region_grafico = st.selectbox("🌎 REGIÓN (Gráfico)", ["Todas"] + sorted(df["REGIÓN"].dropna().unique()), key="region_grafico")
subregion_grafico = st.selectbox("🌏 SUB.REGIÓN (Gráfico)", ["Todas"] + sorted(df["SUB.REGIÓN"].dropna().unique()), key="subregion_grafico")
locacion_grafico = st.selectbox("🏢 LOCACIÓN (Gráfico)", ["Todas"] + sorted(df["LOCACIÓN"].dropna().unique()), key="locacion_grafico")
mesa_grafico = st.selectbox("💼 MESA (Gráfico)", ["Todas"] + sorted(df["MESA"].dropna().unique()), key="mesa_grafico")
ruta_grafico = st.selectbox("🛣️ RUTA (Gráfico)", ["Todas"] + sorted(df["RUTA"].dropna().astype(str).unique()), key="ruta_grafico")

# Filtrar los datos para el gráfico
df_grafico = df_pendientes_total.copy()

if region_grafico != "Todas":
    df_grafico = df_grafico[df_grafico["REGIÓN"] == region_grafico]

if subregion_grafico != "Todas":
    df_grafico = df_grafico[df_grafico["SUB.REGIÓN"] == subregion_grafico]

if locacion_grafico != "Todas":
    df_grafico = df_grafico[df_grafico["LOCACIÓN"] == locacion_grafico]

if mesa_grafico != "Todas":
    df_grafico = df_grafico[df_grafico["MESA"] == mesa_grafico]

if ruta_grafico != "Todas":
    df_grafico = df_grafico[df_grafico["RUTA"].astype(str) == ruta_grafico]

# Agrupar por fecha y contar todos los códigos (permitiendo duplicados en diferentes fechas)
df_chart = df_grafico.groupby("FECHA_ARCHIVO")["CÓDIGO"].count().reset_index()
df_chart = df_chart.rename(columns={"CÓDIGO": "TOTAL_PENDIENTES"}).sort_values("FECHA_ARCHIVO")

# Mostrar gráfico
if not df_chart.empty:
    fig = px.line(df_chart, x="FECHA_ARCHIVO", y="TOTAL_PENDIENTES", markers=True)
    fig.update_layout(
        xaxis_title="Fecha",
        yaxis_title="Total de Pendientes",
        xaxis=dict(tickformat="%d-%m-%Y")
    )
    st.plotly_chart(fig, use_container_width=True)
else:
    st.info("No hay datos suficientes para mostrar el gráfico evolutivo.")

# --- Exportar tabla pendiente del último día ---
def exportar_excel(df_export, nombre):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_export.to_excel(writer, index=False, sheet_name=nombre)
    return output.getvalue()

excel_data1 = exportar_excel(df_ultima_fecha, "PendientesUltimoDia")
st.download_button(
    label="📥 Descargar Excel filtrado",
    data=excel_data1,
    file_name="pendientes_ultimo_dia.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

excel_data2 = exportar_excel(df_pendientes_total, "EvolucionPendientes")
st.download_button(
    label="📥 Descargar evolución de pendientes",
    data=excel_data2,
    file_name="evolucion_pendientes.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
