import streamlit as st
import pandas as pd
import datetime
import plotly.express as px
import io
from openpyxl import Workbook
from openpyxl.styles import PatternFill
from openpyxl.utils.dataframe import dataframe_to_rows

st.set_page_config(page_title="Tablero Operativo - La Virginia", layout="wide", initial_sidebar_state="expanded")

with st.sidebar:
    st.image("https://upload.wikimedia.org/wikipedia/commons/thumb/1/14/Cafe_La_Virginia_logo.svg/1200px-Cafe_La_Virginia_logo.svg.png", width=150)
    st.title("⚙️ Carga de Datos")
    archivo_signum = st.file_uploader("1. Stock (Cygnus)", type=["xlsx", "csv"])
    archivo_sap = st.file_uploader("2. Nómina Materiales (SAP)", type=["xlsx", "csv"])
    st.divider()
    st.caption("Módulo Operativo - Control de Quiebres y Vencimientos")

st.title("☕ Tablero de Control de Stock y Quiebres - La Virginia")

if archivo_signum and archivo_sap:
    try:
        # 1. Carga de datos
        df_signum = pd.read_excel(archivo_signum) if archivo_signum.name.endswith('.xlsx') else pd.read_csv(archivo_signum)
        df_sap = pd.read_excel(archivo_sap) if archivo_sap.name.endswith('.xlsx') else pd.read_csv(archivo_sap)

        df_signum['MATERIAL'] = df_signum['MATERIAL'].astype(str).str.strip()
        df_sap['Material'] = df_sap['Material'].astype(str).str.strip()
        
        # 2. EL CAMBIO CLAVE: Usar SAP como maestro (how='left' desde SAP) para no perder los quiebres invisibles
        df_cruce = pd.merge(df_sap, df_signum, left_on="Material", right_on="MATERIAL", how="left")
        
        # Limpieza de Total (Stock). Si no cruzó con Cygnus, es un Quiebre de stock = 0
        df_cruce['Total'] = pd.to_numeric(df_cruce['Total'], errors='coerce').fillna(0)

        # 3. Fechas
        # Se limpia el .0 que a veces aparece en LOTE al venir vacío
        df_cruce['LOTE_Limpio'] = df_cruce['LOTE'].astype(str).str.replace(r'\.0$', '', regex=True)
        df_cruce['Fecha_Vencimiento'] = pd.to_datetime(df_cruce['LOTE_Limpio'], format='%Y%m%d', errors='coerce')
        hoy = pd.to_datetime(datetime.date.today())
        df_cruce['Dias_Para_Vencer'] = (df_cruce['Fecha_Vencimiento'] - hoy).dt.days

        # 4. Nueva clasificación estricta para Almacén
        def clasificar(row):
            if row['Total'] == 0: return "6. Quiebre (Stock 0)"
            dias = row['Dias_Para_Vencer']
            if pd.isna(dias): return "5. Sin Fecha de Lote"
            if dias < 0: return "1. Vencido"
            elif dias <= 30: return "2. Crítico (<30 días)"
            elif dias <= 60: return "3. Alerta (31-60 días)"
            elif dias <= 90: return "4. Precaución (61-90 días)"
            else: return "0. Óptimo (>90 días)"

        df_cruce['Estado_Alerta'] = df_cruce.apply(clasificar, axis=1)
        df_cruce['Fecha_Mostrada'] = df_cruce['Fecha_Vencimiento'].dt.strftime('%d/%m/%Y').fillna('N/A')

        # 5. Ocupación Física (Pallets)
        if 'UMAxPAI' in df_cruce.columns:
            df_cruce['UMAxPAI'] = pd.to_numeric(df_cruce['UMAxPAI'], errors='coerce').fillna(1)
            df_cruce['Pallets_Ocupados'] = (df_cruce['Total'] / df_cruce['UMAxPAI']).round(2)
        else:
            df_cruce['Pallets_Ocupados'] = 0.0

        # Renombrar columnas para la vista
        if 'NOMBRE' in df_cruce.columns:
            df_cruce.rename(columns={'NOMBRE': 'Descripción'}, inplace=True)
        elif 'NOMBRE_x' in df_cruce.columns:
            df_cruce.rename(columns={'NOMBRE_x': 'Descripción'}, inplace=True)

        # ==========================================
        # PESTAÑAS OPERATIVAS
        # ==========================================
        tab1, tab2, tab3, tab4, tab5 = st.tabs([
            "🚨 CENTRAL DE QUIEBRES", "📊 Vencimientos", "🧊 Ocupación (Pallets)", "🛑 Alerta FEFO (Lotes)", "📋 Reporte Maestro"
        ])

        # --- TAB 1: CENTRAL DE QUIEBRES (LO MÁS IMPORTANTE AHORA) ---
        with tab1:
            st.markdown("### 🚨 Panel de Quiebres y Stock Crítico")
            st.markdown("Monitorea los materiales que se quedaron en cero o están a punto de hacerlo.")
            
            # Botones de filtro rápido
            col_q1, col_q2, col_q3, col_q4 = st.columns(4)
            q_cero = len(df_cruce[df_cruce['Total'] == 0])
            q_10 = len(df_cruce[(df_cruce['Total'] > 0) & (df_cruce['Total'] <= 10)])
            q_50 = len(df_cruce[(df_cruce['Total'] > 10) & (df_cruce['Total'] <= 50)])
            q_100 = len(df_cruce[(df_cruce['Total'] > 50) & (df_cruce['Total'] <= 100)])

            col_q1.metric("🔴 Quiebre Total (Stock 0)", q_cero)
            col_q2.metric("🟠 Peligro (1 a 10 unid.)", q_10)
            col_q3.metric("🟡 Riesgo (11 a 50 unid.)", q_50)
            col_q4.metric("🔵 Alerta Temprana (51 a 100)", q_100)

            st.divider()
            
            tipo_filtro = st.radio("Selecciona qué lista quieres ver para bloquear en SAP:", 
                                   ["Ver Quiebres (Stock = 0)", 
                                    "Ver Stock <= 10", 
                                    "Ver Stock <= 50", 
                                    "Ver Stock <= 100",
                                    "Ver Todo el Riesgo (<= 100)"], horizontal=True)

            if tipo_filtro == "Ver Quiebres (Stock = 0)": df_mostrar = df_cruce[df_cruce['Total'] == 0]
            elif tipo_filtro == "Ver Stock <= 10": df_mostrar = df_cruce[(df_cruce['Total'] > 0) & (df_cruce['Total'] <= 10)]
            elif tipo_filtro == "Ver Stock <= 50": df_mostrar = df_cruce[(df_cruce['Total'] > 10) & (df_cruce['Total'] <= 50)]
            elif tipo_filtro == "Ver Stock <= 100": df_mostrar = df_cruce[(df_cruce['Total'] > 50) & (df_cruce['Total'] <= 100)]
            else: df_mostrar = df_cruce[df_cruce['Total'] <= 100].sort_values(by='Total')

            cols_quiebre = ['Material', 'Descripción', 'Total', 'Pallets_Ocupados']
            st.dataframe(df_mostrar[cols_quiebre], hide_index=True, use_container_width=True)

        # --- TAB 2: VENCIMIENTOS ---
        with tab2:
            st.markdown("### 📊 Estado de Vencimientos")
            k1, k2, k3 = st.columns(3)
            k1.metric("📦 Unidades Totales en Nave", f"{df_cruce['Total'].sum():,.0f}")
            k2.metric("❌ Lotes Vencidos", len(df_cruce[df_cruce['Estado_Alerta'] == '1. Vencido']))
            k3.metric("⚠️ Lotes Críticos (<30 días)", len(df_cruce[df_cruce['Estado_Alerta'] == '2. Crítico (<30 días)']))

            df_pie = df_cruce[~df_cruce['Estado_Alerta'].isin(['5. Sin Fecha de Lote', '6. Quiebre (Stock 0)'])]
            mapa_colores = {"1. Vencido": "#8B0000", "2. Crítico (<30 días)": "#FF0000", "3. Alerta (31-60 días)": "#FFA500", "4. Precaución (61-90 días)": "#FFD700", "0. Óptimo (>90 días)": "#228B22"}
            if not df_pie.empty:
                fig_pie = px.pie(df_pie, names='Estado_Alerta', color='Estado_Alerta', color_discrete_map=mapa_colores, hole=0.4)
                st.plotly_chart(fig_pie, use_container_width=True)

        # --- TAB 3: OCUPACIÓN FÍSICA ---
        with tab3:
            st.markdown("### 🧊 Capacidad Física (Ocupación de Pallets)")
            top_pallets = df_cruce.groupby('Descripción')['Pallets_Ocupados'].sum().reset_index().sort_values(by='Pallets_Ocupados', ascending=False).head(15)
            fig_pallets = px.bar(top_pallets, x='Pallets_Ocupados', y='Descripción', orientation='h', title="Top 15 Productos por Ocupación de Pallets")
            fig_pallets.update_layout(yaxis={'categoryorder':'total ascending'})
            st.plotly_chart(fig_pallets, use_container_width=True)

        # --- TAB 4: ALERTAS FEFO ---
        with tab4:
            st.markdown("### 🛑 Alerta de Mezcla de Lotes (FEFO)")
            st.markdown("Materiales con stock en **más de un lote diferente**. Revisa el almacén para asegurar que se pique el lote más viejo primero.")
            df_activos = df_cruce[df_cruce['Total'] > 0]
            conteo_lotes = df_activos.groupby(['Material', 'Descripción'])['LOTE_Limpio'].nunique().reset_index()
            fefo_alert = conteo_lotes[conteo_lotes['LOTE_Limpio'] > 1].rename(columns={'LOTE_Limpio': 'Cant. Lotes Distintos'})
            st.dataframe(fefo_alert.sort_values(by='Cant. Lotes Distintos', ascending=False), hide_index=True)

        # --- TAB 5: DESCARGA MAESTRA ---
        with tab5:
            st.markdown("### 📋 Base de Datos Completa")
            cols_final = ['Material', 'Descripción', 'LOTE_Limpio', 'Fecha_Mostrada', 'Dias_Para_Vencer', 'Total', 'Pallets_Ocupados', 'Estado_Alerta']
            df_export = df_cruce[cols_final].copy()
            df_export.rename(columns={'LOTE_Limpio': 'Lote Cygnus'}, inplace=True)
            st.dataframe(df_export.head(100), hide_index=True)
            
            def generar_excel_completo(df):
                wb = Workbook()
                ws = wb.active
                ws.title = "Maestro Operativo"
                for r in dataframe_to_rows(df, index=False, header=True):
                    ws.append(r)
                
                # Semaforización en Excel
                fill_gris = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid") # Para Quiebres
                fill_rojo = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
                fill_nar = PatternFill(start_color="FFE6CC", end_color="FFE6CC", fill_type="solid")
                fill_ama = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
                
                idx_estado = len(df.columns)
                for row in ws.iter_rows(min_row=2):
                    estado_val = str(row[idx_estado-1].value)
                    for cell in row:
                        if "6. Quiebre" in estado_val: cell.fill = fill_gris
                        elif "1." in estado_val or "2." in estado_val: cell.fill = fill_rojo
                        elif "3." in estado_val: cell.fill = fill_nar
                        elif "4." in estado_val: cell.fill = fill_ama
                
                output = io.BytesIO()
                wb.save(output)
                return output.getvalue()

            st.download_button("📥 Descargar Excel Operativo", data=generar_excel_completo(df_export), file_name=f"Control_Operativo_La_Virginia_{datetime.date.today()}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    except Exception as e:
        st.error(f"Error procesando datos: {e}")
