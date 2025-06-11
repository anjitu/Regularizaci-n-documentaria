import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO
from datetime import datetime

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

df = cargar_datos()
df["FECHA_ARCHIVO"] = pd.to_datetime(df["FECHA_ARCHIVO"]).dt.date
df = df[df["STATUS A DETALLE"].str.upper() != "COMPLETADO"].copy()

fecha_max = df["FECHA_ARCHIVO"].max()
df_ultimo = df[df["FECHA_ARCHIVO"] == fecha_max]

# Filtros
col1, col2, col3, col4, col5, col6 = st.columns(6)

with col1:
    region = st.selectbox("🌎 REGIÓN", ["Todas"] + sorted(df_ultimo["REGIÓN"].dropna().unique()))
if region != "Todas":
    df_ultimo = df_ultimo[df_ultimo["REGIÓN"] == region]

with col2:
    subregion = st.selectbox("🗺️ SUB.REGIÓN", ["Todas"] + sorted(df_ultimo["SUB.REGIÓN"].dropna().unique()))
if subregion != "Todas":
    df_ultimo = df_ultimo[df_ultimo["SUB.REGIÓN"] == subregion]

with col3:
    locacion = st.selectbox("🏢 LOCACIÓN", ["Todas"] + sorted(df_ultimo["LOCACIÓN"].dropna().unique()))
if locacion != "Todas":
    df_ultimo = df_ultimo[df_ultimo["LOCACIÓN"] == locacion]

with col4:
    mesa = st.selectbox("💼 MESA", ["Todas"] + sorted(df_ultimo["MESA"].dropna().unique()))
if mesa != "Todas":
    df_ultimo = df_ultimo[df_ultimo["MESA"] == mesa]

with col5:
    ruta = st.selectbox("🛣️ RUTA", ["Todas"] + sorted(df_ultimo["RUTA"].dropna().astype(str).unique()))
if ruta != "Todas":
    df_ultimo = df_ultimo[df_ultimo["RUTA"].astype(str) == ruta]

with col6:
    codigo = st.selectbox("🧾 CÓDIGO", ["Todos"] + sorted(df_ultimo["CÓDIGO"].dropna().unique()))
if codigo != "Todos":
    df_ultimo = df_ultimo[df_ultimo["CÓDIGO"] == codigo]

st.markdown(f"🔍 {len(df_ultimo)} pendientes encontrados para el día {fecha_max}")
st.dataframe(df_ultimo, use_container_width=True)

# Evolutivo (con los mismos filtros)
df_filtrado = df.copy()
for filtro, columna in zip([region, subregion, locacion, mesa, ruta, codigo],
                           ["REGIÓN", "SUB.REGIÓN", "LOCACIÓN", "MESA", "RUTA", "CÓDIGO"]):
    if filtro not in ["Todas", "Todos"]:
        df_filtrado = df_filtrado[df_filtrado[columna] == filtro]

df_evolutivo = (
    df_filtrado.groupby("FECHA_ARCHIVO")
    .size().reset_index(name="TOTAL_PENDIENTES")
)

# Gráfico en la app
if not df_evolutivo.empty:
    st.subheader("📈 Evolución de pendientes por fecha")
    fig = px.line(df_evolutivo, x="FECHA_ARCHIVO", y="TOTAL_PENDIENTES", markers=True)
    fig.update_layout(xaxis_title="Fecha", yaxis_title="Total de Pendientes")
    st.plotly_chart(fig, use_container_width=True)
else:
    st.info("No hay datos suficientes para mostrar el gráfico.")

# Función para exportar Excel con formato y gráfico
def exportar_excel(df_data, nombre_hoja, df_grafico=None):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_data.to_excel(writer, index=False, sheet_name=nombre_hoja)
        workbook = writer.book
        worksheet = writer.sheets[nombre_hoja]

        # Encabezado con color
        header_format = workbook.add_format({'bold': True, 'bg_color': '#D9E1F2', 'border': 1})
        for col_num, value in enumerate(df_data.columns.values):
            worksheet.write(0, col_num, value, header_format)
            worksheet.set_column(col_num, col_num, 20)

        # Agregar gráfico si se provee
        if df_grafico is not None and not df_grafico.empty:
            df_grafico.to_excel(writer, sheet_name="Gráfico", index=False)
            chart = workbook.add_chart({'type': 'line'})
            chart.add_series({
                'name': 'Total Pendientes',
                'categories': f"='Gráfico'!$A$2:$A${len(df_grafico)+1}",
                'values': f"='Gráfico'!$B$2:$B${len(df_grafico)+1}",
                'marker': {'type': 'circle'},
            })
            chart.set_title({'name': 'Evolución de Pendientes'})
            chart.set_x_axis({'name': 'Fecha'})
            chart.set_y_axis({'name': 'Total'})
            writer.sheets["Gráfico"].insert_chart("D2", chart)
    return output.getvalue()

# Botones de descarga
excel_1 = exportar_excel(df_ultimo, "Pendientes")
st.download_button("📥 Descargar Excel filtrado", data=excel_1, file_name="pendientes.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

excel_2 = exportar_excel(df_evolutivo, "Evolución", df_evolutivo)
st.download_button("📥 Descargar evolución de pendientes", data=excel_2, file_name="evolucion_pendientes.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")