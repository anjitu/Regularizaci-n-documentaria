import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO
from datetime import datetime

st.set_page_config(page_title="Reporte de Pendientes", layout="wide")
st.title("📊 Reporte de Pendientes de Regularización Documentaria")

# 1. Cargar archivos
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
        fecha_str = archivo.split("-")[-1].replace(".xlsx", "")
        fecha = datetime.strptime(fecha_str, "%d.%m.%Y").date()
        df = pd.read_excel(archivo, sheet_name="BASE TOTAL", dtype=str)
        df["FECHA_ARCHIVO"] = fecha
        df["ARCHIVO_ORIGEN"] = archivo
        dfs.append(df)
    return pd.concat(dfs, ignore_index=True)

df = cargar_datos()

# 2. Filtrar solo pendientes
df["STATUS A DETALLE"] = df["STATUS A DETALLE"].str.upper().str.strip()
df_pendientes = df[df["STATUS A DETALLE"] != "COMPLETADO"].copy()

# 3. Filtros
col1, col2, col3, col4, col5, col6 = st.columns(6)

with col1:
    region = st.selectbox("🌎 REGIÓN", [""] + sorted(df["REGIÓN"].dropna().unique()))
with col2:
    subregion = st.selectbox("🗺️ SUB.REGIÓN", [""] + sorted(df["SUB.REGIÓN"].dropna().unique()))
with col3:
    locacion = st.selectbox("🏢 LOCACIÓN", [""] + sorted(df["LOCACIÓN"].dropna().unique()))
with col4:
    mesa = st.selectbox("🧮 MESA", [""] + sorted(df["MESA"].dropna().unique()))
with col5:
    ruta = st.selectbox("🛣️ RUTA", [""] + sorted(df["RUTA"].dropna().astype(str).unique()))
with col6:
    codigo = st.selectbox("🧾 CÓDIGO", [""] + sorted(df["CÓDIGO"].dropna().astype(str).unique()))

# 4. Aplicar filtros
df_filtrado = df.copy()
if region:
    df_filtrado = df_filtrado[df_filtrado["REGIÓN"] == region]
if subregion:
    df_filtrado = df_filtrado[df_filtrado["SUB.REGIÓN"] == subregion]
if locacion:
    df_filtrado = df_filtrado[df_filtrado["LOCACIÓN"] == locacion]
if mesa:
    df_filtrado = df_filtrado[df_filtrado["MESA"] == mesa]
if ruta:
    df_filtrado = df_filtrado[df_filtrado["RUTA"].astype(str) == ruta]
if codigo:
    df_filtrado = df_filtrado[df_filtrado["CÓDIGO"].astype(str) == codigo]

# 5. Mostrar tabla de resultados filtrados
df_filtrado_pendientes = df_filtrado[df_filtrado["STATUS A DETALLE"] != "COMPLETADO"].copy()
st.markdown(f"🔍 {len(df_filtrado_pendientes)} pendientes encontrados")
st.dataframe(df_filtrado_pendientes, use_container_width=True)

# 6. Evolución de pendientes
df_evolutivo = df_filtrado_pendientes.groupby(
    ["FECHA_ARCHIVO", "REGIÓN", "SUB.REGIÓN", "LOCACIÓN", "MESA", "RUTA"]
).size().reset_index(name="TOTAL_PENDIENTES")

# Mostrar gráfico evolutivo
if not df_evolutivo.empty:
    st.subheader("📈 Evolución de pendientes por fecha")
    fig = px.line(
        df_evolutivo.groupby("FECHA_ARCHIVO")["TOTAL_PENDIENTES"].sum().reset_index(),
        x="FECHA_ARCHIVO",
        y="TOTAL_PENDIENTES",
        markers=True
    )
    fig.update_layout(
        xaxis_title="Fecha",
        yaxis_title="Total de Pendientes",
        xaxis_tickformat="%d-%m-%Y"
    )
    st.plotly_chart(fig, use_container_width=True)
else:
    st.info("No hay datos suficientes para mostrar el gráfico.")

# 7. Función para exportar Excel
def generar_excel(df_export, nombre_hoja):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df_export.to_excel(writer, index=False, sheet_name=nombre_hoja)
        hoja = writer.sheets[nombre_hoja]
        for i, col in enumerate(df_export.columns):
            ancho = max(df_export[col].astype(str).map(len).max(), len(col)) + 2
            hoja.set_column(i, i, ancho)
    return output.getvalue()

# 8. Botón para descargar datos filtrados
excel_data = generar_excel(df_filtrado_pendientes, "Pendientes Filtrados")
st.download_button(
    label="📥 Descargar Excel filtrado",
    data=excel_data,
    file_name="pendientes_filtrados.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

# 9. Botón para descargar evolución
excel_evolutivo = generar_excel(df_evolutivo, "Evolución")
st.download_button(
    label="📊 Descargar evolución de pendientes",
    data=excel_evolutivo,
    file_name="evolucion_pendientes.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)