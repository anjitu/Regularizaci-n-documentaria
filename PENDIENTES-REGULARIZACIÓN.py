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

df_pendientes_total = df[df["STATUS A DETALLE"] != "COMPLETADO"].copy()
fecha_max = df["FECHA_ARCHIVO"].max()
df_ultima_fecha = df_pendientes_total[df_pendientes_total["FECHA_ARCHIVO"] == fecha_max].copy()

# Filtros en cascada para la tabla de pendientes
region = st.selectbox("🌎 REGIÓN", ["Todas"] + sorted(df["REGIÓN"].dropna().unique()), key="region")
if region != "Todas":
    df_ultima_fecha = df_ultima_fecha[df_ultima_fecha["REGIÓN"] == region]

subregion_opts = df[df["REGIÓN"] == region]["SUB.REGIÓN"] if region != "Todas" else df["SUB.REGIÓN"]
subregion = st.selectbox("🌏 SUB.REGIÓN", ["Todas"] + sorted(subregion_opts.dropna().unique()), key="subregion")
if subregion != "Todas":
    df_ultima_fecha = df_ultima_fecha[df_ultima_fecha["SUB.REGIÓN"] == subregion]

locacion_opts = df[df["SUB.REGIÓN"] == subregion]["LOCACIÓN"] if subregion != "Todas" else df["LOCACIÓN"]
locacion = st.selectbox("🏢 LOCACIÓN", ["Todas"] + sorted(locacion_opts.dropna().unique()), key="locacion")
if locacion != "Todas":
    df_ultima_fecha = df_ultima_fecha[df_ultima_fecha["LOCACIÓN"] == locacion]

mesa_opts = df[df["LOCACIÓN"] == locacion]["MESA"] if locacion != "Todas" else df["MESA"]
mesa = st.selectbox("💼 MESA", ["Todas"] + sorted(mesa_opts.dropna().unique()), key="mesa")
if mesa != "Todas":
    df_ultima_fecha = df_ultima_fecha[df_ultima_fecha["MESA"] == mesa]

ruta_opts = df[df["MESA"] == mesa]["RUTA"] if mesa != "Todas" else df["RUTA"]
ruta = st.selectbox("🛣️ RUTA", ["Todas"] + sorted(ruta_opts.dropna().astype(str).unique()), key="ruta")
if ruta != "Todas":
    df_ultima_fecha = df_ultima_fecha[df_ultima_fecha["RUTA"].astype(str) == ruta]

st.markdown(f"🔍 {df_ultima_fecha.shape[0]} pendientes encontrados (fecha {fecha_max})")
st.dataframe(df_ultima_fecha, use_container_width=True)

st.subheader("📈 Evolución de pendientes filtrable")

# Filtros en cascada para el gráfico evolutivo
df_grafico = df_pendientes_total.copy()
if region != "Todas":
    df_grafico = df_grafico[df_grafico["REGIÓN"] == region]
if subregion != "Todas":
    df_grafico = df_grafico[df_grafico["SUB.REGIÓN"] == subregion]
if locacion != "Todas":
    df_grafico = df_grafico[df_grafico["LOCACIÓN"] == locacion]
if mesa != "Todas":
    df_grafico = df_grafico[df_grafico["MESA"] == mesa]
if ruta != "Todas":
    df_grafico = df_grafico[df_grafico["RUTA"].astype(str) == ruta]

# Agrupar solo por fecha
df_evol = df_grafico.groupby("FECHA_ARCHIVO").size().reset_index(name="TOTAL_PENDIENTES")
df_evol["FECHA_ARCHIVO"] = pd.to_datetime(df_evol["FECHA_ARCHIVO"]).dt.strftime("%d/%m/%Y")

if not df_evol.empty:
    fig = px.line(df_evol, x="FECHA_ARCHIVO", y="TOTAL_PENDIENTES", markers=True,
                  title="Evolución de pendientes por fecha")
    fig.update_layout(xaxis_title="Fecha", yaxis_title="Total de Pendientes")
    st.plotly_chart(fig, use_container_width=True)
else:
    st.info("No hay datos suficientes para mostrar el gráfico evolutivo.")

# Excel export
def exportar_excel(df_export, nombre):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df_export.to_excel(writer, index=False, sheet_name=nombre)
    return output.getvalue()

excel_data1 = exportar_excel(df_ultima_fecha, "PendientesUltimoDia")
st.download_button("📥 Descargar Excel de Pendientes Último Día", excel_data1,
                   "pendientes_ultimo_dia.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

excel_data2 = exportar_excel(df_evol, "EvolucionPendientes")
st.download_button("📥 Descargar Excel de Evolución de Pendientes", excel_data2,
                   "evolucion_pendientes.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
