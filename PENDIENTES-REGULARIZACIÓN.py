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

region = st.selectbox("🌎 REGIÓN", ["Todas"] + sorted(df["REGIÓN"].dropna().unique()), key="region")
if region != "Todas":
    df_ultima_fecha = df_ultima_fecha[df_ultima_fecha["REGIÓN"] == region]

subregion_options = ["Todas"] + sorted(df[df["REGIÓN"] == region]["SUB.REGIÓN"].dropna().unique()) if region != "Todas" else ["Todas"] + sorted(df["SUB.REGIÓN"].dropna().unique())
subregion = st.selectbox("🌏 SUB.REGIÓN", subregion_options, key="subregion")
if subregion != "Todas":
    df_ultima_fecha = df_ultima_fecha[df_ultima_fecha["SUB.REGIÓN"] == subregion]

locaciones = df["LOCACIÓN"].dropna().unique()
locacion = st.selectbox("🏢 LOCACIÓN", ["Todas"] + sorted(locaciones), key="locacion")
if locacion != "Todas":
    df_ultima_fecha = df_ultima_fecha[df_ultima_fecha["LOCACIÓN"] == locacion]

mesas = df["MESA"].dropna().unique()
mesa = st.selectbox("💼 MESA", ["Todas"] + sorted(mesas), key="mesa")
if mesa != "Todas":
    df_ultima_fecha = df_ultima_fecha[df_ultima_fecha["MESA"] == mesa]

rutas = df["RUTA"].dropna().astype(str).unique()
ruta = st.selectbox("🛣️ RUTA", ["Todas"] + sorted(rutas), key="ruta")
if ruta != "Todas":
    df_ultima_fecha = df_ultima_fecha[df_ultima_fecha["RUTA"].astype(str) == ruta]

codigos = df_ultima_fecha["CÓDIGO"].dropna().unique()
codigo = st.selectbox("🔢 CÓDIGO", ["Todos"] + sorted(codigos), key="codigo")
if codigo != "Todos":
    df_ultima_fecha = df_ultima_fecha[df_ultima_fecha["CÓDIGO"] == codigo]

st.markdown(f"🔍 {df_ultima_fecha.shape[0]} pendientes encontrados (fecha {fecha_max})")
st.dataframe(df_ultima_fecha, use_container_width=True)

st.subheader("📈 Evolución de pendientes filtrable")

region_g = st.selectbox("🌎 REGIÓN (Gráfico)", ["Todas"] + sorted(df["REGIÓN"].dropna().unique()), key="region_g")
if region_g != "Todas":
    df_aux = df_pendientes_total[df_pendientes_total["REGIÓN"] == region_g]
else:
    df_aux = df_pendientes_total.copy()

subregion_g_opt = ["Todas"] + sorted(df[df["REGIÓN"] == region_g]["SUB.REGIÓN"].dropna().unique()) if region_g != "Todas" else ["Todas"] + sorted(df["SUB.REGIÓN"].dropna().unique())
subregion_g = st.selectbox("🌏 SUB.REGIÓN (Gráfico)", subregion_g_opt, key="subregion_g")
if subregion_g != "Todas":
    df_aux = df_aux[df_aux["SUB.REGIÓN"] == subregion_g]

locacion_g_opt = ["Todas"] + sorted(df["LOCACIÓN"].dropna().unique())
locacion_g = st.selectbox("🏢 LOCACIÓN (Gráfico)", locacion_g_opt, key="locacion_g")
if locacion_g != "Todas":
    df_aux = df_aux[df_aux["LOCACIÓN"] == locacion_g]

mesa_g_opt = ["Todas"] + sorted(df["MESA"].dropna().unique())
mesa_g = st.selectbox("💼 MESA (Gráfico)", mesa_g_opt, key="mesa_g")
if mesa_g != "Todas":
    df_aux = df_aux[df_aux["MESA"] == mesa_g]

ruta_g_opt = ["Todas"] + sorted(df["RUTA"].dropna().astype(str).unique())
ruta_g = st.selectbox("🛣️ RUTA (Gráfico)", ruta_g_opt, key="ruta_g")
if ruta_g != "Todas":
    df_aux = df_aux[df_aux["RUTA"].astype(str) == ruta_g]

df_evol = df_aux.groupby(
    ["REGIÓN", "SUB.REGIÓN", "LOCACIÓN", "MESA", "RUTA", "FECHA_ARCHIVO"]
).size().reset_index(name="TOTAL_PENDIENTES")

df_pivot = df_evol.pivot_table(
    index=["REGIÓN", "SUB.REGIÓN", "LOCACIÓN", "MESA", "RUTA"],
    columns="FECHA_ARCHIVO",
    values="TOTAL_PENDIENTES",
    fill_value=0
).reset_index()

df_pivot.columns.name = None
df_pivot.columns = [col.strftime("%d/%m/%Y") if isinstance(col, (pd.Timestamp, datetime)) else col for col in df_pivot.columns]

st.dataframe(df_pivot, use_container_width=True)

df_melt = df_pivot.melt(
    id_vars=["REGIÓN", "SUB.REGIÓN", "LOCACIÓN", "MESA", "RUTA"],
    var_name="Fecha",
    value_name="Total Pendientes"
)

if not df_melt.empty:
    fig = px.line(
        df_melt,
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

excel_data2 = exportar_excel(df_pivot, "EvolucionPendientes")
st.download_button(
    label="📥 Descargar Excel de Evolución de Pendientes",
    data=excel_data2,
    file_name="evolucion_pendientes.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
