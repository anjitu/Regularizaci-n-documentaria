import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO
from datetime import datetime

st.set_page_config(page_title="Reporte de Pendientes", layout="wide")
st.title("Reporte de Pendientes de Regularización Documentaria")

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
df_pendientes = df[df["STATUS A DETALLE"].str.upper() != "COMPLETADO"].copy()

# --- 3. Filtros dependientes ---
col1, col2, col3, col4, col5, col6 = st.columns(6)

with col1:
    region_opciones = [""] + sorted(df_pendientes["REGIÓN"].dropna().unique())
    region = st.selectbox("🌎 REGIÓN", region_opciones)

df_subreg = df_pendientes[df_pendientes["REGIÓN"] == region] if region else df_pendientes
with col2:
    subregion_opciones = [""] + sorted(df_subreg["SUB.REGIÓN"].dropna().unique())
    subregion = st.selectbox("🗺️ SUB.REGIÓN", subregion_opciones)

df_loc = df_subreg[df_subreg["SUB.REGIÓN"] == subregion] if subregion else df_subreg
with col3:
    locacion_opciones = [""] + sorted(df_loc["LOCACIÓN"].dropna().unique())
    locacion = st.selectbox("🏢 LOCACIÓN", locacion_opciones)

df_mesa = df_loc[df_loc["LOCACIÓN"] == locacion] if locacion else df_loc
with col4:
    mesa_opciones = [""] + sorted(df_mesa["MESA"].dropna().unique())
    mesa = st.selectbox("MESA", mesa_opciones)

df_ruta = df_mesa[df_mesa["MESA"] == mesa] if mesa else df_mesa
with col5:
    ruta_opciones = [""] + sorted(df_ruta["RUTA"].dropna().astype(str).unique())
    ruta = st.selectbox("🛣️ RUTA", ruta_opciones)

with col6:
    codigo_cliente_opciones = [""] + sorted(df_ruta["CÓDIGO DE CLIENTE"].dropna().astype(str).unique())
    codigo_cliente = st.selectbox("🧾 CÓDIGO DE CLIENTE", codigo_cliente_opciones)

# --- 4. Aplicar filtros ---
df_filtrado = df_ruta.copy()
if ruta:
    df_filtrado = df_filtrado[df_filtrado["RUTA"].astype(str) == ruta]
if codigo_cliente:
    df_filtrado = df_filtrado[df_filtrado["CÓDIGO DE CLIENTE"].astype(str) == codigo_cliente]

# --- 5. Mostrar resultados ---
st.markdown(f"🔍 {len(df_filtrado)} resultados encontrados")
st.dataframe(df_filtrado, use_container_width=True)

# --- 6. Crear resumen dinámico según filtros aplicados ---
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

df_evolutivo = df_resumen.groupby("FECHA_ARCHIVO").size().reset_index(name="TOTAL_PENDIENTES")

# --- 7. Mostrar gráfico evolutivo ---
if not df_evolutivo.empty:
    st.subheader("📈 Evolución de pendientes en el tiempo (según filtros)")
    fig = px.line(df_evolutivo, x="FECHA_ARCHIVO", y="TOTAL_PENDIENTES",
                  markers=True, title="Tendencia de pendientes diarios")
    fig.update_layout(xaxis_title="Fecha", yaxis_title="N° de Pendientes")
    st.plotly_chart(fig, use_container_width=True)
else:
    st.info("No hay datos suficientes para mostrar un gráfico evolutivo con los filtros actuales.")

# --- 8. Función para exportar Excel bonito ---
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
            if column_name.upper() == "STATUS A DETALLE":
                worksheet.write(0, col_num, column_name, yellow_header_format)
            else:
                worksheet.write(0, col_num, column_name, header_format)

        for i, col in enumerate(df_export.columns):
            col_width = max(df_export[col].astype(str).map(len).max(), len(col)) + 2
            worksheet.set_column(i, i, col_width)

        worksheet.freeze_panes(1, 0)
        worksheet.autofilter(0, 0, df_export.shape[0], df_export.shape[1] - 1)

    return output.getvalue()

# --- 9. Botón para descargar Excel ---
excel_data = to_excel_bytes(df_filtrado)
st.download_button(
    label="📥 Descargar Excel filtrado",
    data=excel_data,
    file_name="pendientes_filtrados.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
