import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime
from io import BytesIO

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
df["STATUS A DETALLE"] = df["STATUS A DETALLE"].str.upper()

# Filtramos solo pendientes
df_pendientes_total = df[df["STATUS A DETALLE"] != "COMPLETADO"].copy()

# Último día
fecha_max = df["FECHA_ARCHIVO"].max()
df_ultima_fecha = df_pendientes_total[df_pendientes_total["FECHA_ARCHIVO"] == fecha_max].copy()

# FILTROS EN CASCADA
region = st.selectbox("🌎 REGIÓN", ["Todas"] + sorted(df["REGIÓN"].dropna().unique()), key="region")
if region != "Todas":
    df_ultima_fecha = df_ultima_fecha[df_ultima_fecha["REGIÓN"] == region]
    df_pendientes_total = df_pendientes_total[df_pendientes_total["REGIÓN"] == region]

subregiones = df_pendientes_total["SUB.REGIÓN"].dropna().unique()
subregion = st.selectbox("🌏 SUB.REGIÓN", ["Todas"] + sorted(subregiones), key="subregion")
if subregion != "Todas":
    df_ultima_fecha = df_ultima_fecha[df_ultima_fecha["SUB.REGIÓN"] == subregion]
    df_pendientes_total = df_pendientes_total[df_pendientes_total["SUB.REGIÓN"] == subregion]

locaciones = df_pendientes_total["LOCACIÓN"].dropna().unique()
locacion = st.selectbox("🏢 LOCACIÓN", ["Todas"] + sorted(locaciones), key="locacion")
if locacion != "Todas":
    df_ultima_fecha = df_ultima_fecha[df_ultima_fecha["LOCACIÓN"] == locacion]
    df_pendientes_total = df_pendientes_total[df_pendientes_total["LOCACIÓN"] == locacion]

mesas = df_pendientes_total["MESA"].dropna().unique()
mesa = st.selectbox("MESA", ["Todas"] + sorted(mesas), key="mesa")
if mesa != "Todas":
    df_ultima_fecha = df_ultima_fecha[df_ultima_fecha["MESA"] == mesa]
    df_pendientes_total = df_pendientes_total[df_pendientes_total["MESA"] == mesa]

rutas = df_pendientes_total["RUTA"].dropna().astype(str).unique()
ruta = st.selectbox("🛣️ RUTA", ["Todas"] + sorted(rutas), key="ruta")
if ruta != "Todas":
    df_ultima_fecha = df_ultima_fecha[df_ultima_fecha["RUTA"].astype(str) == ruta]
    df_pendientes_total = df_pendientes_total[df_pendientes_total["RUTA"].astype(str) == ruta]

# Mostrar tabla de últimos pendientes
st.markdown(f"🔍 {df_ultima_fecha.shape[0]} pendientes encontrados (fecha {fecha_max})")
st.dataframe(df_ultima_fecha, use_container_width=True)

# 📈 Evolución de pendientes por fecha
st.subheader("📈 Evolución de pendientes por fecha")

df_evol = df_pendientes_total.groupby(
    ["REGIÓN", "SUB.REGIÓN", "LOCACIÓN", "MESA", "RUTA", "FECHA_ARCHIVO"]
).size().reset_index(name="TOTAL_PENDIENTES")

# Tabla dinámica estilo tabla pivote
pivot = df_evol.pivot_table(
    index=["REGIÓN", "SUB.REGIÓN", "LOCACIÓN", "MESA", "RUTA"],
    columns="FECHA_ARCHIVO",
    values="TOTAL_PENDIENTES",
    fill_value=0
).reset_index()

pivot.columns.name = None
pivot.columns = [col.strftime("%d/%m/%Y") if isinstance(col, (pd.Timestamp, datetime)) else col for col in pivot.columns]
st.dataframe(pivot, use_container_width=True)

# Gráfico evolutivo
melt = pivot.melt(
    id_vars=["REGIÓN", "SUB.REGIÓN", "LOCACIÓN", "MESA", "RUTA"],
    var_name="Fecha",
    value_name="Total Pendientes"
)

if not melt.empty:
    fig = px.line(
        melt,
        x="Fecha",
        y="Total Pendientes",
        color="REGIÓN",
        line_group="RUTA",
        markers=True,
        title="Evolución de pendientes por fecha"
    )
    fig.update_layout(xaxis_title="Fecha", yaxis_title="Total de Pendientes")
    st.plotly_chart(fig, use_container_width=True)
else:
    st.info("No hay datos suficientes para mostrar el gráfico evolutivo.")

# Botones de descarga

def exportar_excel(df_export, nombre):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df_export.to_excel(writer, index=False, sheet_name=nombre)
    return output.getvalue()

excel_data1 = exportar_excel(df_ultima_fecha, "PendientesUltimoDia")
st.download_button(
    label="📥 Descargar Excel de Pendientes Último Día",
    data=excel_data1,
    file_name="pendientes_ultimo_dia.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

excel_data2 = exportar_excel(pivot, "EvolucionPendientes")
st.download_button(
    label="📥 Descargar Excel de Evolución de Pendientes",
    data=excel_data2,
    file_name="evolucion_pendientes.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
