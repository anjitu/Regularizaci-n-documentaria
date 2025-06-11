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

# --- Crear tabla dinámica para la evolución de pendientes ---
# Filtrar solo las fechas de interés
fechas_interes = df_pendientes_total["FECHA_ARCHIVO"].unique()
df_pendientes_total = df_pendientes_total[df_pendientes_total["FECHA_ARCHIVO"].isin(fechas_interes)]

# Contar pendientes por fecha, región, subregión, locación, mesa y ruta
df_evolucion = df_pendientes_total.groupby(
    ["FECHA_ARCHIVO", "REGIÓN", "SUB.REGIÓN", "LOCACIÓN", "MESA", "RUTA"]
).size().reset_index(name="TOTAL_PENDIENTES")

# Pivotar la tabla para tener las fechas como columnas
df_pivot = df_evolucion.pivot_table(
    index=["REGIÓN", "SUB.REGIÓN", "LOCACIÓN", "MESA", "RUTA"],
    columns="FECHA_ARCHIVO",
    values="TOTAL_PENDIENTES",
    fill_value=0
).reset_index()

# Renombrar las columnas de fecha
df_pivot.columns.name = None  # Eliminar el nombre de la columna
df_pivot = df_pivot.rename(columns={df_pivot.columns[i]: df_pivot.columns[i].strftime("%d/%m/%Y") for i in range(1, len(df_pivot.columns))})

# Mostrar la tabla de evolución
st.subheader("📊 Evolución de Pendientes")
st.dataframe(df_pivot, use_container_width=True)

# --- Mostrar gráfico de evolución ---
df_chart = df_pivot.melt(id_vars=["REGIÓN", "SUB.REGIÓN", "LOCACIÓN", "MESA", "RUTA"], var_name="FECHA", value_name="TOTAL_PENDIENTES")

# Graficar
if not df_chart.empty:
    fig = px.line(df_chart, x="FECHA", y="TOTAL_PENDIENTES", color="REGIÓN", line_group="RUTA", markers=True)
    fig.update_layout(
        xaxis_title="Fecha",
        yaxis_title="Total de Pendientes",
        xaxis=dict(tickformat="%d-%m-%Y"),
        title="Evolución de Pendientes por Fecha"
    )
    st.plotly_chart(fig, use_container_width=True)
else:
    st.info("No hay datos suficientes para mostrar el gráfico evolutivo.")

# --- Exportar tabla de evolución ---
def exportar_excel(df_export, nombre):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_export.to_excel(writer, index=False, sheet_name=nombre)
    return output.getvalue()

excel_data = exportar_excel(df_pivot, "EvolucionPendientes")
st.download_button(
    label="📥 Descargar evolución de pendientes",
    data=excel_data,
    file_name="evolucion_pendientes.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
