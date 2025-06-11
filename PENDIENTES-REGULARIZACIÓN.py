import streamlit as st
import pandas as pd
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
mesa = st.selectbox("💼 MESA", ["Todas"] + sorted(mesas), key="mesa")
if mesa != "Todas":
    df_ultima_fecha = df_ultima_fecha[df_ultima_fecha["MESA"] == mesa]
    df_pendientes_total = df_pendientes_total[df_pendientes_total["MESA"] == mesa]

rutas = df_pendientes_total["RUTA"].dropna().astype(str).unique()
ruta = st.selectbox("🛣️ RUTA", ["Todas"] + sorted(rutas), key="ruta")
if ruta != "Todas":
    df_ultima_fecha = df_ultima_fecha[df_ultima_fecha["RUTA"].astype(str) == ruta]
    df_pendientes_total = df_pendientes_total[df_pendientes_total["RUTA"].astype(str) == ruta]

# Filtro por CÓDIGO
codigos = df_pendientes_total["CÓDIGO"].dropna().unique()
codigo = st.selectbox("🧾 CÓDIGO DE CLIENTE", ["Todos"] + sorted(codigos), key="codigo")
if codigo != "Todos":
    df_ultima_fecha = df_ultima_fecha[df_ultima_fecha["CÓDIGO"] == codigo]

# Mostrar tabla de últimos pendientes
st.markdown(f"🔍 {df_ultima_fecha.shape[0]} pendientes encontrados (fecha {fecha_max})")
st.dataframe(df_ultima_fecha, use_container_width=True)

# 📊 Matriz de evolución de pendientes
st.subheader("🧮 Matriz de Evolución de Pendientes por Fecha")

df_evol = df_pendientes_total.groupby(
    ["REGIÓN", "SUB.REGIÓN", "LOCACIÓN", "MESA", "RUTA", "FECHA_ARCHIVO"]
).size().reset_index(name="TOTAL_PENDIENTES")

# Crear matriz
pivot = df_evol.pivot_table(
    index=["REGIÓN", "SUB.REGIÓN", "LOCACIÓN", "MESA", "RUTA"],
    columns="FECHA_ARCHIVO",
    values="TOTAL_PENDIENTES",
    fill_value=0
)

# Ordenar columnas por fecha
pivot = pivot.sort_index(axis=1)

# Reiniciar índice y renombrar columnas
pivot = pivot.reset_index()
pivot.columns.name = None
pivot.columns = [col.strftime("%d/%m/%Y") if isinstance(col, (pd.Timestamp, datetime, datetime.date)) else col for col in pivot.columns]

# Mostrar matriz
st.dataframe(pivot, use_container_width=True)

# Función para exportar Excel con formato
def exportar_excel(df_export, nombre):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df_export.to_excel(writer, index=False, sheet_name=nombre)

        workbook = writer.book
        worksheet = writer.sheets[nombre]

        # Estilo encabezado
        header_format = workbook.add_format({
            "bold": True,
            "text_wrap": True,
            "valign": "center",
            "fg_color": "#D9E1F2",
            "border": 1
        })

        for col_num, column in enumerate(df_export.columns):
            col_values = df_export[column].astype(str)
            max_len = max(col_values.map(len).max(), len(str(column))) + 2
            worksheet.set_column(col_num, col_num, max_len)
            worksheet.write(0, col_num, str(column), header_format)

    return output.getvalue()

# Botones de descarga
excel_data1 = exportar_excel(df_ultima_fecha, "PendientesUltimoDia")
st.download_button(
    label="📥 Descargar Excel de Pendientes Último Día",
    data=excel_data1,
    file_name="pendientes_ultimo_dia.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

excel_data2 = exportar_excel(pivot, "MatrizPendientes")
st.download_button(
    label="📥 Descargar Excel de Matriz de Evolución",
    data=excel_data2,
    file_name="matriz_evolucion_pendientes.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
