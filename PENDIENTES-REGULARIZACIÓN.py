import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO
from datetime import datetime

st.set_page_config(page_title="Reporte de Pendientes", layout="wide")
st.title("📊 Reporte de Pendientes de Regularización Documentaria")

# --- Cargar archivos automáticamente ---
@st.cache_data
def cargar_datos():
    archivos = [
        "CEO-LISTA DE PENDIENTES-09.06.2025.xlsx",
        "NORTE-LISTA DE PENDIENTES-09.06.2025.xlsx",
        "LIMA-LISTA DE PENDIENTES-09.06.2025.xlsx",
        "SUR-LISTA DE PENDIENTES-09.06.2025.xlsx"
    ]
    dfs = []
    for archivo in archivos:
        try:
            fecha_str = archivo.split("-")[-1].replace(".xlsx", "")
            fecha = datetime.strptime(fecha_str, "%d.%m.%Y").date()
        except:
            fecha = None
        df = pd.read_excel(archivo, sheet_name="BASE TOTAL", dtype=str)
        df["FECHA_ARCHIVO"] = fecha
        df["ARCHIVO_ORIGEN"] = archivo
        dfs.append(df)
    return pd.concat(dfs, ignore_index=True)

df = cargar_datos()

# --- Filtrar pendientes (STATUS distinto a "COMPLETADO") ---
df["STATUS A DETALLE"] = df["STATUS A DETALLE"].fillna("").str.upper()
df_pendientes = df[df["STATUS A DETALLE"] != "COMPLETADO"].copy()

# --- Filtros dinámicos ---
st.markdown("### 🎛️ Filtros")
col1, col2, col3, col4, col5, col6 = st.columns(6)

with col1:
    region = st.selectbox("🌎 REGIÓN", [""] + sorted(df_pendientes["REGIÓN"].dropna().unique()))
df_filtered = df_pendientes[df_pendientes["REGIÓN"] == region] if region else df_pendientes

with col2:
    subregion = st.selectbox("🗺️ SUB.REGIÓN", [""] + sorted(df_filtered["SUB.REGIÓN"].dropna().unique()))
df_filtered = df_filtered[df_filtered["SUB.REGIÓN"] == subregion] if subregion else df_filtered

with col3:
    locacion = st.selectbox("🏢 LOCACIÓN", [""] + sorted(df_filtered["LOCACIÓN"].dropna().unique()))
df_filtered = df_filtered[df_filtered["LOCACIÓN"] == locacion] if locacion else df_filtered

with col4:
    mesa = st.selectbox("MESA", [""] + sorted(df_filtered["MESA"].dropna().unique()))
df_filtered = df_filtered[df_filtered["MESA"] == mesa] if mesa else df_filtered

with col5:
    ruta = st.selectbox("🛣️ RUTA", [""] + sorted(df_filtered["RUTA"].dropna().astype(str).unique()))
df_filtered = df_filtered[df_filtered["RUTA"].astype(str) == ruta] if ruta else df_filtered

with col6:
    codigo = st.selectbox("🧾 CÓDIGO", [""] + sorted(df_filtered["CÓDIGO"].dropna().astype(str).unique()))
df_filtered = df_filtered[df_filtered["CÓDIGO"].astype(str) == codigo] if codigo else df_filtered

# --- Mostrar resultados filtrados ---
st.markdown(f"🔍 {len(df_filtered)} resultados encontrados")
st.dataframe(df_filtered, use_container_width=True)

# --- Gráfico Evolutivo según filtros ---
df_evol = df_filtered.groupby("FECHA_ARCHIVO").size().reset_index(name="TOTAL_PENDIENTES")
df_evol["FECHA_ARCHIVO"] = pd.to_datetime(df_evol["FECHA_ARCHIVO"])

if not df_evol.empty:
    fig = px.line(df_evol, x="FECHA_ARCHIVO", y="TOTAL_PENDIENTES", markers=True,
                  title="📈 Evolución de Pendientes por Día")
    fig.update_layout(xaxis_title="Fecha", yaxis_title="N° de Pendientes")
    fig.update_traces(line_color="#C00000")
    fig.update_xaxes(dtick="D1", tickformat="%d-%m")
    st.plotly_chart(fig, use_container_width=True)
else:
    st.warning("No hay datos suficientes para mostrar el gráfico evolutivo.")

# --- Descargar Excel con resumen por fecha, región, etc. ---
df_descarga = df_filtered.groupby([
    "FECHA_ARCHIVO", "REGIÓN", "SUB.REGIÓN", "LOCACIÓN", "MESA", "RUTA"
]).size().reset_index(name="TOTAL_PENDIENTES")

def to_excel_bytes(df_export):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_export.to_excel(writer, index=False, sheet_name='Evolucion')
        ws = writer.sheets['Evolucion']
        for i, col in enumerate(df_export.columns):
            col_width = max(df_export[col].astype(str).map(len).max(), len(col)) + 2
            ws.set_column(i, i, col_width)
        ws.freeze_panes(1, 0)
    return output.getvalue()

st.download_button(
    label="📥 Descargar resumen evolutivo (Excel)",
    data=to_excel_bytes(df_descarga),
    file_name="evolucion_pendientes.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)