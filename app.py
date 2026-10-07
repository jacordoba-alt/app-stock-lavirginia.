import streamlit as st
import pandas as pd
import datetime
import plotly.express as px
import io
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font
from openpyxl.utils.dataframe import dataframe_to_rows

st.set_page_config(page_title="Control de Stock - La Virginia", layout="wide")
st.title("☕ Dashboard de Stock y Vencimientos - La Virginia")
st.markdown("Sube los reportes de **Signum (Stock)** y **SAP (Materiales)** para ver el estado al instante.")

col1, col2 = st.columns(2)
with col1:
    archivo_signum = st.file_uploader("1. Subir Archivo de Stock (Signum)", type=["xlsx", "csv"])
with col2:
    archivo_sap = st.file_uploader("2. Subir Nómina de Materiales (SAP)", type=["xlsx", "csv"])

if archivo_signum and archivo_sap:
    # Cargar datos
    df_signum = pd.read_excel(archivo_signum) if archivo_signum.name.endswith('.xlsx') else pd.read_csv(archivo_signum)
    df_sap = pd.read_excel(archivo_sap) if archivo_sap.name.endswith('.xlsx') else pd.read_csv(archivo_sap)

    # ⚠️ CAMBIA ESTOS NOMBRES SI TUS EXCEL TIENEN COLUMNAS DISTINTAS ⚠️
    col_clave = "Codigo_Material" # Columna que une ambos archivos
    col_stock = "Stock"
    col_venc = "Fecha_Vencimiento"
    
    # Cruce
    df_cruce = pd.merge(df_signum, df_sap, on=col_clave, how="left")
    df_cruce[col_venc] = pd.to_datetime(df_cruce[col_venc])
    hoy = pd.to_datetime(datetime.date.today())
    df_cruce['Dias_Para_Vencer'] = (df_cruce[col_venc] - hoy).dt.days

    def clasificar(dias):
        if dias < 0: return "Vencido"
        elif dias <= 30: return "Crítico (<30 días)"
        elif dias <= 60: return "Alerta (<60 días)"
        else: return "Óptimo"

    df_cruce['Estado'] = df_cruce['Dias_Para_Vencer'].apply(clasificar)
    df_cruce[col_venc] = df_cruce[col_venc].dt.strftime('%d/%m/%Y')

    st.success("✅ Datos procesados exitosamente")
    
    # KPIs
    k1, k2, k3 = st.columns(3)
    k1.metric("🔴 Materiales Stock CERO", len(df_cruce[df_cruce[col_stock] == 0]))
    k2.metric("⚠️ Lotes Críticos (<30 días)", len(df_cruce[df_cruce['Estado'].str.contains('Crítico|Vencido')]))
    k3.metric("📦 Stock Total", df_cruce[col_stock].sum())

    # Gráfico
    st.subheader("Distribución de Vencimientos")
    fig = px.pie(df_cruce, names='Estado', color='Estado', 
                 color_discrete_map={"Vencido": "darkred", "Crítico (<30 días)": "red", "Alerta (<60 días)": "orange", "Óptimo": "green"})
    st.plotly_chart(fig, use_container_width=True)

    # Botón para descargar Excel
    def generar_excel(df):
        wb = Workbook()
        ws = wb.active
        ws.title = "Reporte"
        for r in dataframe_to_rows(df, index=False, header=True):
            ws.append(r)
        
        # Colores
        fill_rojo = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
        fill_ama = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
        fill_ver = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
        
        # Asumiendo que la columna de 'Estado' queda al final
        idx_estado = len(df.columns)
        for row in ws.iter_rows(min_row=2):
            estado_val = row[idx_estado-1].value
            for cell in row:
                if estado_val in ["Vencido", "Crítico (<30 días)"]: cell.fill = fill_rojo
                elif estado_val == "Alerta (<60 días)": cell.fill = fill_ama
                else: cell.fill = fill_ver
                
        output = io.BytesIO()
        wb.save(output)
        return output.getvalue()

    st.download_button(
        label="📥 Descargar Excel Semaforizado",
        data=generar_excel(df_cruce),
        file_name=f"Reporte_Stock_{datetime.date.today()}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
else:
    st.info("👆 Esperando archivos...")
