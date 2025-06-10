import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO
from datetime import datetime

st.set_page_config(page_title="Reporte de Pendientes", layout="wide")
st.title("📋 Reporte de Pendientes de Regularización Documentaria")

# --- 1. Carga automática de archivos con fecha extraída del nombre ---
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
        try:
            fecha = datetime.strptime(fecha_str, "%d.%m.%Y").date()
        except:
            fecha = None
        df = pd.read_excel(archivo, sheet_name="BASE TOTAL", dtype=str)
        df["ARCHIVO_ORIGEN"] = archivo
        df["FECHA_ARCHIVO"] = fecha
        dfs.append(df)
    return pd.concat(dfs, ignore_index=True)

df = cargar_datos()

# --- 2. Filtrar solo pendientes ---
df["STATUS A DETALLE"] = df["STATUS A DETALLE"].fillna("")
df_pendientes = df[df["STATUS A DETALLE"].str.upper() != "COMPLETADO"].copy()

# --- 3. Filtros dependientes ---
col1, col2, col3, col4, col5, col6 = st.columns(6)

with col1:
    region_opciones = [""] + sorted(df_pendientes["REGIÓN"].dropna().unique())
    region = st.selectbox("🌎 REGIÓN", region_opciones)

f1 = df_pendientes[df_pendientes["REGIÓN"] == region] if region else df_pendientes
with col2:
    subregion_opciones = [""] + sorted(f1["SUB.REGIÓN"].dropna().unique())
    subregion = st.selectbox("🗺️ SUB.REGIÓN", subregion_opciones)

f2 = f1[f1["SUB.REGIÓN"] == subregion] if subregion else f1
with col3:
    locacion_opciones = [""] + sorted(f2["LOCACIÓN"].dropna().unique())
    locacion = st.selectbox("🏢 LOCACIÓN", locacion_opciones)

f3 = f2[f2["LOCACIÓN"] == locacion] if locacion else f2
with col4:
    mesa_opciones = [""] + sorted(f3["MESA"].dropna().unique())
    mesa = st.selectbox("🧾 MESA", mesa_opciones)

f4 = f3[f3["MESA"] == mesa] if mesa else f3
with col5:
    ruta_opciones = [""] + sorted(f4["RUTA"].dropna().astype(str).unique())
    ruta = st.selectbox("🛣️ RUTA", ruta_opciones)

with col6:
    codigo_cliente_opciones = [""] + sorted(f4["CÓDIGO"].dropna().astype(str).unique())
    codigo_cliente = st.selectbox("📌 CÓDIGO CLIENTE", codigo_cliente_opciones)

# --- 4. Aplicar filtros finales ---
df_filtrado = f4.copy()
if ruta:
    df_filtrado = df_filtrado[df_filtrado["RUTA"].astype(str) == ruta]
if codigo_cliente:
    df_filtrado = df_filtrado[df_filtrado["CÓDIGO"].astype(str) == codigo_cliente]

st.markdown(f"🔍 {len(df_filtrado)} resultados encontrados")
st.dataframe(df_filtrado, use_container_width=True)

# --- 5. Gráfico evolutivo ---
df_resumen = df_pendientes.copy()
if region:
    df_resumen = df_resumen[df_resumen["REGIÓN"] == region]
if subregion:
    df_resumen = df_resumen[df_resumen["SUB.REGIÓN"] == subregion]
if locacion:
    df_resumen = df_resumen[df_resumen["LOCACIÓN"] == locacion]
if mesa:
    df_resumen = df_resumen[df_resumen["MESA"] == mesa]
if ruta:
    df_resumen = df_resumen[df_resumen["RUTA"].astype(str) == ruta]

# Asegurar formato correcto de fecha
df_resumen["FECHA_ARCHIVO"] = pd.to_datetime(df_resumen["FECHA_ARCHIVO"]).dt.date

df_evolutivo = (
    df_resumen.groupby("FECHA_ARCHIVO").size().reset_index(name="TOTAL_PENDIENTES")
)

if not df_evolutivo.empty:
    st.subheader("📈 Evolución de pendientes (según filtros)")
    fig = px.line(
        df_evolutivo,
        x="FECHA_ARCHIVO",
        y="TOTAL_PENDIENTES",
        markers=True,
        title="Tendencia de pendientes diarios"
    )
    fig.update_layout(
        xaxis_title="Fecha",
        yaxis_title="N° de Pendientes",
        xaxis=dict(tickformat="%d-%m-%Y"),
        yaxis=dict(range=[0, df_evolutivo["TOTAL_PENDIENTES"].max() * 1.2])
    )
    st.plotly_chart(fig, use_container_width=True)
else:
    st.info("No hay datos suficientes para mostrar un gráfico evolutivo con los filtros actuales.")

# --- 6. Descargar Excel filtrado (clientes) ---
def to_excel_bytes(df_export):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_export.to_excel(writer, index=False, sheet_name='Pendientes')
        workbook = writer.book
        worksheet = writer.sheets['Pendientes']

        header_format = workbook.add_format({
            'bold': True,
            'bg_color': '#C00000',
            'font_color': 'white',
            'border': 1
        })
        yellow_header_format = workbook.add_format({
            'bold': True,
            'bg_color': '#FFF2CC',
            'border': 1
        })

        for col_num, column_name in enumerate(df_export.columns):
            fmt = yellow_header_format if column_name.upper() == "STATUS A DETALLE" else header_format
            worksheet.write(0, col_num, column_name, fmt)
            col_width = max(df_export[column_name].astype(str).map(len).max(), len(column_name)) + 2
            worksheet.set_column(col_num, col_num, col_width)

        worksheet.freeze_panes(1, 0)
        worksheet.autofilter(0, 0, df_export.shape[0], df_export.shape[1] - 1)
    return output.getvalue()

# --- Botón descarga Excel (tabla de datos filtrada) ---
excel_data = to_excel_bytes(df_filtrado)
st.download_button(
    label="📥 Descargar Excel filtrado",
    data=excel_data,
    file_name="pendientes_filtrados.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)