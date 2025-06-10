import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO
from datetime import datetime

st.set_page_config(page_title="📊 Reporte de Pendientes", layout="wide")
st.title("📋 Reporte de Pendientes de Regularización Documentaria")

# 📂 Carga el archivo evolutivo (pendientes históricos)
@st.cache_data
def cargar_evolutivo():
    try:
        return pd.read_excel("evolutivo_pendientes.xlsx", dtype=str, parse_dates=["FECHA_ARCHIVO"])
    except:
        return pd.DataFrame()

df_evolutivo_full = cargar_evolutivo()

# 📋 Filtros en cascada
col1, col2, col3, col4, col5, col6 = st.columns(6)

with col1:
    region = st.selectbox("🌎 REGIÓN", [""] + sorted(df_evolutivo_full["REGIÓN"].dropna().unique()))
df_subreg = df_evolutivo_full[df_evolutivo_full["REGIÓN"] == region] if region else df_evolutivo_full

with col2:
    subregion = st.selectbox("🗺️ SUB.REGIÓN", [""] + sorted(df_subreg["SUB.REGIÓN"].dropna().unique()))
df_loc = df_subreg[df_subreg["SUB.REGIÓN"] == subregion] if subregion else df_subreg

with col3:
    locacion = st.selectbox("🏢 LOCACIÓN", [""] + sorted(df_loc["LOCACIÓN"].dropna().unique()))
df_mesa = df_loc[df_loc["LOCACIÓN"] == locacion] if locacion else df_loc

with col4:
    mesa = st.selectbox("MESA", [""] + sorted(df_mesa["MESA"].dropna().unique()))
df_ruta = df_mesa[df_mesa["MESA"] == mesa] if mesa else df_mesa

with col5:
    ruta = st.selectbox("🛣️ RUTA", [""] + sorted(df_ruta["RUTA"].dropna().astype(str).unique()))
df_codigo = df_ruta[df_ruta["RUTA"] == ruta] if ruta else df_ruta

with col6:
    codigo = st.selectbox("🧾 CÓDIGO", [""] + sorted(df_codigo["CÓDIGO"].dropna().astype(str).unique()))
df_filtrado = df_codigo[df_codigo["CÓDIGO"] == codigo] if codigo else df_codigo

# 📊 Mostrar gráfico de evolución
df_grafico = df_filtrado.groupby("FECHA_ARCHIVO").agg({"TOTAL_PENDIENTES": "sum"}).reset_index()

if not df_grafico.empty:
    st.subheader("📈 Evolución de pendientes en el tiempo")
    fig = px.line(df_grafico, x="FECHA_ARCHIVO", y="TOTAL_PENDIENTES", markers=True,
                  title="📉 Tendencia de Pendientes", text="TOTAL_PENDIENTES")
    fig.update_traces(textposition="top center")
    fig.update_layout(xaxis_title="📅 Fecha", yaxis_title="🔢 Total de Pendientes", yaxis_range=[0, df_grafico["TOTAL_PENDIENTES"].max() + 5])
    fig.update_xaxes(dtick="D", tickformat="%d-%m-%Y")
    st.plotly_chart(fig, use_container_width=True)
else:
    st.warning("⚠️ No hay datos para mostrar el gráfico con los filtros actuales.")

# 📎 Botón para descargar Excel filtrado
def to_excel_bytes(df_export):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df_export.to_excel(writer, index=False, sheet_name="Pendientes")
        workbook = writer.book
        worksheet = writer.sheets["Pendientes"]
        header_format = workbook.add_format({
            'bold': True,
            'bg_color': '#C00000',
            'font_color': 'white',
            'border': 1
        })
        for col_num, column_name in enumerate(df_export.columns):
            worksheet.write(0, col_num, column_name, header_format)
            col_width = max(df_export[column_name].astype(str).map(len).max(), len(column_name)) + 2
            worksheet.set_column(col_num, col_num, col_width)
        worksheet.freeze_panes(1, 0)
        worksheet.autofilter(0, 0, df_export.shape[0], df_export.shape[1] - 1)
    return output.getvalue()

# 📥 Botón 1: Descargar pendientes filtrados
if not df_filtrado.empty:
    excel_filtrado = to_excel_bytes(df_filtrado)
    st.download_button(
        label="📥 Descargar pendientes filtrados",
        data=excel_filtrado,
        file_name="pendientes_filtrados.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# 📥 Botón 2: Descargar resumen evolutivo
if not df_grafico.empty:
    df_resumen_excel = df_filtrado[["FECHA_ARCHIVO", "REGIÓN", "SUB.REGIÓN", "LOCACIÓN", "MESA", "RUTA", "TOTAL_PENDIENTES"]]
    df_resumen_excel = df_resumen_excel.groupby(["FECHA_ARCHIVO", "REGIÓN", "SUB.REGIÓN", "LOCACIÓN", "MESA", "RUTA"]).agg({"TOTAL_PENDIENTES": "sum"}).reset_index()
    resumen_bytes = to_excel_bytes(df_resumen_excel)
    st.download_button(
        label="📤 Descargar evolución de pendientes",
        data=resumen_bytes,
        file_name="evolucion_pendientes.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )