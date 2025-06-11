import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO
from datetime import datetime
import xlsxwriter

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
        try:
            fecha = datetime.strptime(fecha_str, "%d.%m.%Y").date()
        except:
            fecha = None
        df = pd.read_excel(archivo, sheet_name="BASE TOTAL", dtype=str)
        df["ARCHIVO_ORIGEN"] = archivo
        df["FECHA_ARCHIVO"] = fecha
        dfs.append(df)
    return pd.concat(dfs, ignore_index=True)

# --- Carga y preparación de datos ---
df = cargar_datos()
df["FECHA_ARCHIVO"] = pd.to_datetime(df["FECHA_ARCHIVO"]).dt.date
df["STATUS A DETALLE"] = df["STATUS A DETALLE"].str.upper()

# --- Pendientes para todo el histórico (gráfico) y para último día (tabla) ---
df_todos_pendientes = df[df["STATUS A DETALLE"] != "COMPLETADO"].copy()
ultima_fecha = df["FECHA_ARCHIVO"].max()
df_ultimo_dia = df_todos_pendientes[df_todos_pendientes["FECHA_ARCHIVO"] == ultima_fecha].copy()

# --- Filtros (afectan tanto a tabla como gráfico) ---
col1, col2, col3, col4, col5, col6 = st.columns(6)

with col1:
    region = st.selectbox("🌎 REGIÓN", ["Todas"] + sorted(df_todos_pendientes["REGIÓN"].dropna().unique()))
with col2:
    subregion = st.selectbox("🗺️ SUB.REGIÓN", ["Todas"] + sorted(df_todos_pendientes["SUB.REGIÓN"].dropna().unique()))
with col3:
    locacion = st.selectbox("🏢 LOCACIÓN", ["Todas"] + sorted(df_todos_pendientes["LOCACIÓN"].dropna().unique()))
with col4:
    mesa = st.selectbox("MESA", ["Todas"] + sorted(df_todos_pendientes["MESA"].dropna().unique()))
with col5:
    ruta = st.selectbox("🛣️ RUTA", ["Todas"] + sorted(df_todos_pendientes["RUTA"].dropna().astype(str).unique()))
with col6:
    codigo = st.selectbox("🔢 CÓDIGO", ["Todos"] + sorted(df_todos_pendientes["CÓDIGO"].dropna().unique()))

# --- Aplicar filtros ---
def aplicar_filtros(df_base):
    if region != "Todas":
        df_base = df_base[df_base["REGIÓN"] == region]
    if subregion != "Todas":
        df_base = df_base[df_base["SUB.REGIÓN"] == subregion]
    if locacion != "Todas":
        df_base = df_base[df_base["LOCACIÓN"] == locacion]
    if mesa != "Todas":
        df_base = df_base[df_base["MESA"] == mesa]
    if ruta != "Todas":
        df_base = df_base[df_base["RUTA"].astype(str) == ruta]
    if codigo != "Todos":
        df_base = df_base[df_base["CÓDIGO"] == codigo]
    return df_base

df_filtrado_tabla = aplicar_filtros(df_ultimo_dia)
df_filtrado_grafico = aplicar_filtros(df_todos_pendientes)

# --- Mostrar tabla filtrada ---
st.markdown(f"🔍 *{len(df_filtrado_tabla)}* pendientes encontrados")
st.dataframe(df_filtrado_tabla, use_container_width=True)

# --- Evolución de pendientes (por CÓDIGO) ---
df_evolutivo = (
    df_filtrado_grafico.groupby(["FECHA_ARCHIVO", "REGIÓN", "SUB.REGIÓN", "LOCACIÓN", "MESA", "RUTA"])
    .agg(TOTAL_PENDIENTES=('CÓDIGO', 'count')).reset_index()
)

# --- Mostrar gráfico dinámico ---
if not df_evolutivo.empty:
    st.subheader("📈 Evolución de pendientes por fecha")
    df_chart = df_evolutivo.groupby("FECHA_ARCHIVO")["TOTAL_PENDIENTES"].sum().reset_index()
    fig = px.line(df_chart, x="FECHA_ARCHIVO", y="TOTAL_PENDIENTES", markers=True)
    fig.update_layout(xaxis_title="Fecha", yaxis_title="Total de Pendientes")
    st.plotly_chart(fig, use_container_width=True)
else:
    st.info("No hay datos suficientes para mostrar el gráfico evolutivo.")

# --- Exportar Excel con formato y gráfico ---
def exportar_excel(df_export, df_evol):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df_export.to_excel(writer, sheet_name="Pendientes", index=False)
        df_evol.to_excel(writer, sheet_name="Evolución", index=False)

        workbook = writer.book
        worksheet = writer.sheets["Pendientes"]
        header_format = workbook.add_format({'bold': True, 'bg_color': '#DCE6F1'})
        for col_num, value in enumerate(df_export.columns.values):
            worksheet.write(0, col_num, value, header_format)
            max_len = max(df_export[value].astype(str).map(len).max(), len(value)) + 2
            worksheet.set_column(col_num, col_num, max_len)

        evol_ws = writer.sheets["Evolución"]
        for col_num, value in enumerate(df_evol.columns.values):
            evol_ws.write(0, col_num, value, header_format)
            max_len = max(df_evol[value].astype(str).map(len).max(), len(value)) + 2
            evol_ws.set_column(col_num, col_num, max_len)

        if not df_chart.empty:
            chart = workbook.add_chart({'type': 'line'})
            chart.add_series({
                'name': 'Total Pendientes',
                'categories': ['Evolución', 1, 0, len(df_chart), 0],
                'values':     ['Evolución', 1, 6, len(df_chart), 6],
                'marker': {'type': 'circle'},
            })
            chart.set_title({'name': 'Evolución de Pendientes'})
            chart.set_x_axis({'name': 'Fecha'})
            chart.set_y_axis({'name': 'Pendientes'})
            evol_ws.insert_chart('H2', chart)

    return output.getvalue()

# --- Botones de descarga ---
excel_data = exportar_excel(df_filtrado_tabla, df_evolutivo)

st.download_button(
    label="📥 Descargar Excel completo",
    data=excel_data,
    file_name="reporte_pendientes.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)