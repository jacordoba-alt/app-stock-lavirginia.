import streamlit as st
import pandas as pd
import datetime
import plotly.express as px
import io
from openpyxl import Workbook
from openpyxl.styles import PatternFill
from openpyxl.utils.dataframe import dataframe_to_rows

st.set_page_config(page_title="Control de Stock - La Virginia", layout="wide")
st.title("☕ Dashboard de Stock y Vencimientos - La Virginia")
st.markdown("Sube los reportes de **Cygnus (Stock)** y **SAP (Materiales)** para ver el estado al instante.")

col1, col2 = st.columns(2)
with col1:
    archivo_signum = st.file_uploader("1. Subir Archivo de Stock (Cygnus)", type=["xlsx", "csv"])
with col2:
    archivo_sap = st.file_uploader("2. Subir Nómina de Materiales (SAP)", type=["xlsx", "csv"])

if archivo_signum and archivo_sap:
    try:
        # Cargar datos
        df_signum = pd.read_excel(archivo_signum) if archivo_signum.name.endswith('.xlsx') else pd.read_csv(archivo_signum)
        df_sap = pd.read_excel(archivo_sap) if archivo_sap.name.endswith('.xlsx') else pd.read_csv(archivo_sap)

        # 1. Emparejar las columnas clave (limpiando espacios y convirtiendo a texto)
        df_signum['MATERIAL'] = df_signum['MATERIAL'].astype(str).str.strip()
        df_sap['Material'] = df_sap['Material'].astype(str).str.strip()

        # 2. Hacer el cruce uniendo 'MATERIAL' de Cygnus con 'Material' de SAP
        df_cruce = pd.merge(df_signum, df_sap, left_on="MATERIAL", right_on="Material", how="left")
        
        # 3. Convertir el LOTE (formato YYYYMMDD) a una fecha real que la máquina entienda
        df_cruce['Fecha_Vencimiento'] = pd.to_datetime(df_cruce['LOTE'].astype(str), format='%Y%m%d', errors='coerce')
        
        # 4. Calcular días faltantes
        hoy = pd.to_datetime(datetime.date.today())
        df_cruce['Dias_Para_Vencer'] = (df_cruce['Fecha_Vencimiento'] - hoy).dt.days

        def clasificar(dias):
            if pd.isna(dias): return "Sin Fecha"
            if dias < 0: return "Vencido"
            elif dias <= 30: return "Crítico (<30 días)"
            elif dias <= 60: return "Alerta (<60 días)"
            else: return "Óptimo"

        df_cruce['Estado'] = df_cruce['Dias_Para_Vencer'].apply(clasificar)
        
        # Formatear la fecha para que en el Excel final se vea bonita (Día/Mes/Año)
        df_cruce['Fecha_Mostrada'] = df_cruce['Fecha_Vencimiento'].dt.strftime('%d/%m/%Y').fillna('Sin Fecha')

        st.success("✅ Datos procesados y cruzados exitosamente")
        
        # 5. Configurar KPIs (Total es la columna de stock en Cygnus)
        col_stock = "Total"
        # Asegurarnos de que el Total sea un número
        df_cruce[col_stock] = pd.to_numeric(df_cruce[col_stock], errors='coerce').fillna(0)
        
        k1, k2, k3 = st.columns(3)
        k1.metric("🔴 Lotes con Stock CERO", len(df_cruce[df_cruce[col_stock] == 0]))
        k2.metric("⚠️ Lotes Críticos (<30 días)", len(df_cruce[df_cruce['Estado'].str.contains('Crítico|Vencido', na=False)]))
        k3.metric("📦 Unidades Totales de Stock", f"{df_cruce[col_stock].sum():,.0f}")

        # 6. Gráfico
        st.subheader("Distribución de Vencimientos")
        df_grafico = df_cruce[df_cruce['Estado'] != 'Sin Fecha'] # Ocultamos los que no tienen fecha del gráfico
        if not df_grafico.empty:
            fig = px.pie(df_grafico, names='Estado', color='Estado', 
                         color_discrete_map={"Vencido": "darkred", "Crítico (<30 días)": "red", "Alerta (<60 días)": "orange", "Óptimo": "green"})
            st.plotly_chart(fig, use_container_width=True)

        # 7. Botón para descargar Excel ordenado
        def generar_excel(df):
            # Nos quedamos solo con las columnas más importantes para el vendedor
            columnas_finales = ['MATERIAL', 'NOMBRE_x', 'LOTE', 'Fecha_Mostrada', 'Dias_Para_Vencer', 'Total', 'Estado']
            # Filtramos para evitar errores si falta alguna columna
            cols_existentes = [c for c in columnas_finales if c in df.columns]
            df_export = df[cols_existentes].copy()
            # Renombramos "NOMBRE_x" a "Descripción" para que quede más limpio
            if 'NOMBRE_x' in df_export.columns:
                df_export.rename(columns={'NOMBRE_x': 'Descripción'}, inplace=True)
            
            wb = Workbook()
            ws = wb.active
            ws.title = "Reporte Cruzado"
            for r in dataframe_to_rows(df_export, index=False, header=True):
                ws.append(r)
            
            # Aplicar los colores en el Excel
            fill_rojo = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
            fill_ama = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
            fill_ver = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
            
            idx_estado = len(cols_existentes)
            for row in ws.iter_rows(min_row=2):
                estado_val = row[idx_estado-1].value
                for cell in row:
                    if estado_val in ["Vencido", "Crítico (<30 días)"]: cell.fill = fill_rojo
                    elif estado_val == "Alerta (<60 días)": cell.fill = fill_ama
                    elif estado_val == "Óptimo": cell.fill = fill_ver
                    
            output = io.BytesIO()
            wb.save(output)
            return output.getvalue()

        st.download_button(
            label="📥 Descargar Excel Semaforizado",
            data=generar_excel(df_cruce),
            file_name=f"Reporte_Stock_Vencimientos_{datetime.date.today()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"Hubo un error procesando los datos. Detalle: {e}")
else:
    st.info("👆 Por favor sube ambos archivos (Cygnus y SAP) para generar el reporte.")
