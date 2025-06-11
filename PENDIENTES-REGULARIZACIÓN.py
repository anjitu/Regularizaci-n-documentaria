import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO
from datetime import datetime

st.set_page_config(page_title="Reporte de Pendientes", layout="wide")
st.title("📋 Consulta de Pendientes de Regularización Documentaria")

# --- 1. Cargar múltiples archivos ---
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

# Cargar datos
original_df = cargar_datos()
original_df["FECHA_ARCHIVO"] = pd.to_datetime(original_df["FECHA_ARCHIVO"]).dt.date

# Crear copia con todos los pendientes históricos
df_historico_pendientes = original_df[original_df["STATUS A DETALLE"].str.upper() != "COMPLETADO"].copy()

# Obtener última fecha cargada
ultima_fecha = max(original_df["FECHA_ARCHIVO"])

# Filtrar para mostrar solo los pendientes de la última fecha en la tabla principal
df_pendientes = df_historico_pendientes[df_historico_pendientes["FECHA_ARCHIVO"] == ultima_fecha].copy()

# --- 2. Filtros dependientes ---
col1, col2, col3, col4, col5, col6 = st.columns(6)

with col1:
    region = st.selectbox("🌎 REGIÓN", ["Todas"] + sorted(df_pendientes["REGIÓN"].dropna().unique()))
if region != "Todas":
    df_pendientes = df_pendientes[df_pendientes["REGIÓN"] == region]

with col2:
    subregion = st.selectbox("🗺️ SUB.REGIÓN", ["Todas"] + sorted(df_pendientes["SUB.REGIÓN"].dropna().unique()))
if subregion != "Todas":
    df_pendientes = df_pendientes[df_pendientes["SUB.REGIÓN"] == subregion]

with col3:
    locacion = st.selectbox("🏢 LOCACIÓN", ["Todas"] + sorted(df_pendientes["LOCACIÓN"].dropna().unique()))
if locacion != "Todas":
    df_pendientes = df_pendientes[df_pendientes["LOCACIÓN"] == locacion]

with col4:
    mesa = st.selectbox("MESA", ["Todas"] + sorted(df_pendientes["MESA"].dropna().unique()))
if mesa != "Todas":
    df_pendientes = df_pendientes[df_pendientes["MESA"] == mesa]

with col5:
    ruta = st.selectbox("🛣️ RUTA", ["Todas"] + sorted(df_pendientes["RUTA"].dropna().astype(str).unique()))
if ruta != "Todas":
    df_pendientes = df_pendientes[df_pendientes["RUTA"].astype(str) == ruta]

with col6:
    codigo = st.selectbox("🧾 CÓDIGO", ["Todas"] + sorted(df_pendientes["CÓDIGO"].dropna().astype(str).unique()))
if codigo != "Todas":
    df_pendientes = df_pendientes[df_pendientes["CÓDIGO"].astype(str) == codigo]

# --- 3. Mostrar tabla filtrada ---
st.markdown(f"🔍 {len(df_pendientes)} pendientes encontrados para el {ultima_fecha}")
st.dataframe(df_pendientes, use_container_width=True)

# --- 4. Crear evolución de pendientes (con todos los días) ---
df_evolutivo = (
    df_historico_pendientes.groupby([
        "FECHA_ARCHIVO", "REGIÓN", "SUB.REGIÓN", "LOCACIÓN", "MESA", "RUTA"])
    .size().reset_index(name="TOTAL_PENDIENTES")
)

# --- 5. Mostrar gráfico evolutivo ---
if not df_evolutivo.empty:
    st.subheader("📈 Evolución de pendientes por fecha")
    df_chart = df_evolutivo.groupby("FECHA_ARCHIVO")["TOTAL_PENDIENTES"].sum().reset_index()
    fig = px.line(df_chart, x="FECHA_ARCHIVO", y="TOTAL_PENDIENTES", markers=True)
    fig.update_layout(xaxis_title="Fecha", yaxis_title="Total de Pendientes")
    st.plotly_chart(fig, use_container_width=True)
else:
    st.info("No hay datos suficientes para mostrar el gráfico evolutivo.")

# --- 6. Función para exportar a Excel ---
def to_excel_bytes(df_export, nombre_hoja):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_export.to_excel(writer, index=False, sheet_name=nombre_hoja)
    return output.getvalue()

# --- 7. Botón para descargar tabla filtrada ---
st.download_button(
    label="📥 Descargar Excel filtrado",
    data=to_excel_bytes(df_pendientes, "Pendientes"),
    file_name="pendientes_filtrados.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

# --- 8. Botón para descargar evolución ---
st.download_button(
    label="📥 Descargar evolución de pendientes",
    data=to_excel_bytes(df_evolutivo, "EvolucionPendientes"),
    file_name="evolucion_pendientes.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)