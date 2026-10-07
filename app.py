import streamlit as st
import pandas as pd
import datetime
import plotly.express as px
import io
from openpyxl import Workbook
from openpyxl.styles import PatternFill
from openpyxl.utils.dataframe import dataframe_to_rows

# Configuración de página a pantalla completa
st.set_page_config(page_title="Control de Stock - La Virginia", layout="wide", initial_sidebar_state="expanded")

# --- BARRA LATERAL (Para subir archivos y no ensuciar la vista principal) ---
with st.sidebar:
    st.image("https://upload.wikimedia.org/wikipedia/commons/thumb/1/14/Cafe_La_Virginia_logo.svg/1200px-Cafe_La_Virginia_logo.svg.png", width=150)
    st.title("⚙️ Carga de Datos")
    st.markdown("Actualiza los reportes diarios aquí:")
    archivo_signum = st.file_uploader("1. Stock (Cygnus)", type=["xlsx", "csv"])
    archivo_sap = st.file_uploader("2. Nómina Materiales (SAP)", type=["xlsx", "csv"])
    st.divider()
    st.caption("Dashboard Gerencial v2.0 - Área de Operaciones")

# --- PANTALLA PRINCIPAL ---
st.title("☕ Dashboard Gerencial de Stock y Vencimientos")

if archivo_signum and archivo_sap:
    try:
        # Carga de datos
        df_signum = pd.read_excel(archivo_signum) if archivo_signum.name.endswith('.xlsx') else pd.read_csv(archivo_signum)
        df_sap = pd.read_excel(archivo_sap) if archivo_sap.name.endswith('.xlsx') else pd.read_csv(archivo_sap)

        # Limpieza y cruce
        df_signum['MATERIAL'] = df_signum['MATERIAL'].astype(str).str.strip()
        df_sap['Material'] = df_sap['Material'].astype(str).str.strip()
        df_cruce = pd.merge(df_signum, df_sap, left_on="MATERIAL", right_on="Material", how="left")
        
        # Procesamiento de Fechas y Días
        df_cruce['Fecha_Vencimiento'] = pd.to_datetime(df_cruce['LOTE'].astype(str), format='%Y%m%d', errors='coerce')
        hoy = pd.to_datetime(datetime.date.today())
        df_cruce['Dias_Para_Vencer'] = (df_cruce['Fecha_Vencimiento'] - hoy).dt.days

        # Nueva clasificación más detallada (Gerencial)
        def clasificar(dias):
            if pd.isna(dias): return "5. Sin Fecha"
            if dias < 0: return "1. Vencido"
            elif dias <= 30: return "2. Crítico (<30 días)"
            elif dias <= 60: return "3. Alerta (31-60 días)"
            elif dias <= 90: return "4. Precaución (61-90 días)"
            else: return "0. Óptimo (>90 días)"

        df_cruce['Estado_Alerta'] = df_cruce['Dias_Para_Vencer'].apply(clasificar)
        df_cruce['Fecha_Mostrada'] = df_cruce['Fecha_Vencimiento'].dt.strftime('%d/%m/%Y').fillna('Sin Fecha')
        
        # Formateo de Stock
        col_stock = "Total"
        df_cruce[col_stock] = pd.to_numeric(df_cruce[col_stock], errors='coerce').fillna(0)
        
        # Ordenar columnas para las vistas
        columnas_vista = ['MATERIAL', 'NOMBRE_x', 'LOTE', 'Fecha_Mostrada', 'Dias_Para_Vencer', 'Total', 'Estado_Alerta']
        df_final = df_cruce[columnas_vista].copy()
        df_final.rename(columns={'NOMBRE_x': 'Descripción'}, inplace=True)
        df_final = df_final.sort_values(by='Dias_Para_Vencer')

        # --- CREACIÓN DE PESTAÑAS (TABS) ---
        tab1, tab2, tab3 = st.tabs(["📊 Visión Global", "🚨 Panel de Acción (Bloqueos)", "📋 Base de Datos y Descarga"])

        # PESTAÑA 1: VISIÓN GLOBAL (Gráficos y KPIs)
        with tab1:
            st.markdown("### 📈 Indicadores Clave de Rendimiento (KPIs)")
            k1, k2, k3, k4 = st.columns(4)
            k1.metric("📦 Unidades Totales en Stock", f"{df_final['Total'].sum():,.0f}")
            k2.metric("🔴 Materiales Stock CERO", len(df_final[df_final['Total'] == 0]))
            k3.metric("⚠️ Lotes Críticos (<30 días)", len(df_final[df_final['Estado_Alerta'] == '2. Crítico (<30 días)']))
            k4.metric("🟡 Lotes en Precaución (<90 días)", len(df_final[df_final['Estado_Alerta'].str.contains('3|4', na=False)]))

            st.divider()
            col_graf1, col_graf2 = st.columns(2)
            
            with col_graf1:
                st.markdown("**Distribución de Riesgo de Vencimiento**")
                df_pie = df_final[df_final['Estado_Alerta'] != '5. Sin Fecha']
                mapa_colores = {
                    "1. Vencido": "#8B0000", 
                    "2. Crítico (<30 días)": "#FF0000", 
                    "3. Alerta (31-60 días)": "#FFA500", 
                    "4. Precaución (61-90 días)": "#FFD700",
                    "0. Óptimo (>90 días)": "#228B22"
                }
                if not df_pie.empty:
                    fig_pie = px.pie(df_pie, names='Estado_Alerta', color='Estado_Alerta', color_discrete_map=mapa_colores, hole=0.4)
                    st.plotly_chart(fig_pie, use_container_width=True)

            with col_graf2:
                st.markdown("**Top 10 Materiales con Mayor Volumen de Stock**")
                top_stock = df_final.groupby('Descripción')['Total'].sum().reset_index().sort_values(by='Total', ascending=False).head(10)
                fig_bar = px.bar(top_stock, x='Total', y='Descripción', orientation='h', color_discrete_sequence=['#1F497D'])
                fig_bar.update_layout(yaxis={'categoryorder':'total ascending'})
                st.plotly_chart(fig_bar, use_container_width=True)

        # PESTAÑA 2: PANEL DE ACCIÓN (Para bloqueos en SAP)
        with tab2:
            st.markdown("### 🛑 Gestión de Quiebres de Stock y Bloqueos")
            st.markdown("Utiliza esta sección para identificar rápidamente los productos que deben ser bloqueados para la venta.")
            
            col_acc1, col_acc2 = st.columns([1, 2])
            with col_acc1:
                st.info("Ajusta el parámetro para ver materiales con bajo stock:")
                umbral_stock = st.slider("Mostrar materiales con stock menor o igual a:", min_value=0, max_value=500, value=50, step=10)
            
            with col_acc2:
                df_bloqueo = df_final[df_final['Total'] <= umbral_stock].sort_values(by='Total')
                st.error(f"Se encontraron **{len(df_bloqueo)}** lotes con {umbral_stock} unidades o menos.")
                st.dataframe(df_bloqueo, use_container_width=True, hide_index=True)

        # PESTAÑA 3: BASE DE DATOS Y EXCEL
        with tab3:
            st.markdown("### 📋 Vista Detallada de Inventario")
            
            # Función para colorear la tabla en la web
            def color_filas_web(row):
                estado = row['Estado_Alerta']
                if "1." in estado or "2." in estado: return ['background-color: #ffcccc; color: black'] * len(row)
                elif "3." in estado: return ['background-color: #ffe6cc; color: black'] * len(row)
                elif "4." in estado: return ['background-color: #fff2cc; color: black'] * len(row)
                elif "0." in estado: return ['background-color: #e6f2ff; color: black'] * len(row)
                return [''] * len(row)

            st.dataframe(df_final.style.apply(color_filas_web, axis=1), use_container_width=True, hide_index=True)

            # Botón de Descarga Excel
            def generar_excel_gerencial(df):
                wb = Workbook()
                ws = wb.active
                ws.title = "Reporte Gerencial"
                for r in dataframe_to_rows(df, index=False, header=True):
                    ws.append(r)
                
                fill_rojo = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
                fill_nar = PatternFill(start_color="FFE6CC", end_color="FFE6CC", fill_type="solid")
                fill_ama = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
                fill_azul = PatternFill(start_color="E6F2FF", end_color="E6F2FF", fill_type="solid")
                
                for row in ws.iter_rows(min_row=2):
                    estado_val = str(row[6].value) # Columna Estado_Alerta
                    for cell in row:
                        if "1." in estado_val or "2." in estado_val: cell.fill = fill_rojo
                        elif "3." in estado_val: cell.fill = fill_nar
                        elif "4." in estado_val: cell.fill = fill_ama
                        elif "0." in estado_val: cell.fill = fill_azul
                        
                output = io.BytesIO()
                wb.save(output)
                return output.getvalue()

            st.download_button(
                label="📥 Descargar Reporte Completo en Excel",
                data=generar_excel_gerencial(df_final),
                file_name=f"Reporte_Gerencial_La_Virginia_{datetime.date.today()}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Hubo un error procesando los datos. Detalle: {e}")
else:
    st.info("👆 Por favor sube ambos archivos en la barra lateral izquierda para generar el reporte gerencial.")
