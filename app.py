import streamlit as st
import pandas as pd
import datetime
import plotly.express as px
import io
from openpyxl import Workbook
from openpyxl.styles import PatternFill
from openpyxl.utils.dataframe import dataframe_to_rows

st.set_page_config(page_title="Tablero Estratégico - La Virginia", layout="wide", initial_sidebar_state="expanded")

with st.sidebar:
    st.image("https://upload.wikimedia.org/wikipedia/commons/thumb/1/14/Cafe_La_Virginia_logo.svg/1200px-Cafe_La_Virginia_logo.svg.png", width=150)
    st.title("⚙️ Carga de Datos")
    archivo_signum = st.file_uploader("1. Stock (Cygnus)", type=["xlsx", "csv"])
    archivo_sap = st.file_uploader("2. Nómina Materiales (SAP)", type=["xlsx", "csv"])
    st.divider()
    st.caption("Módulo Gerencial v3.0 - Operaciones")

st.title("☕ Tablero Estratégico Integral - La Virginia")

if archivo_signum and archivo_sap:
    try:
        # 1. Carga de datos
        df_signum = pd.read_excel(archivo_signum) if archivo_signum.name.endswith('.xlsx') else pd.read_csv(archivo_signum)
        df_sap = pd.read_excel(archivo_sap) if archivo_sap.name.endswith('.xlsx') else pd.read_csv(archivo_sap)

        df_signum['MATERIAL'] = df_signum['MATERIAL'].astype(str).str.strip()
        df_sap['Material'] = df_sap['Material'].astype(str).str.strip()
        df_cruce = pd.merge(df_signum, df_sap, left_on="MATERIAL", right_on="Material", how="left")
        
        # 2. Fechas
        df_cruce['Fecha_Vencimiento'] = pd.to_datetime(df_cruce['LOTE'].astype(str), format='%Y%m%d', errors='coerce')
        hoy = pd.to_datetime(datetime.date.today())
        df_cruce['Dias_Para_Vencer'] = (df_cruce['Fecha_Vencimiento'] - hoy).dt.days

        def clasificar(dias):
            if pd.isna(dias): return "5. Sin Fecha"
            if dias < 0: return "1. Vencido"
            elif dias <= 30: return "2. Crítico (<30 días)"
            elif dias <= 60: return "3. Alerta (31-60 días)"
            elif dias <= 90: return "4. Precaución (61-90 días)"
            else: return "0. Óptimo (>90 días)"

        df_cruce['Estado_Alerta'] = df_cruce['Dias_Para_Vencer'].apply(clasificar)
        df_cruce['Fecha_Mostrada'] = df_cruce['Fecha_Vencimiento'].dt.strftime('%d/%m/%Y').fillna('Sin Fecha')
        df_cruce['Total'] = pd.to_numeric(df_cruce['Total'], errors='coerce').fillna(0)

        # ==========================================
        # LOS 4 NUEVOS MOTORES ESTRATÉGICOS
        # ==========================================
        
        # MOTOR 1: Ocupación Física (Pallets)
        if 'UMAxPAI' in df_cruce.columns:
            df_cruce['UMAxPAI'] = pd.to_numeric(df_cruce['UMAxPAI'], errors='coerce').fillna(1)
            df_cruce['Pallets_Ocupados'] = (df_cruce['Total'] / df_cruce['UMAxPAI']).round(2)
        else:
            df_cruce['Pallets_Ocupados'] = 0.0

        # MOTOR 2: Simulación Inteligente (Si faltan columnas de Precio y Venta en SAP)
        simulado = False
        if 'Precio' not in df_cruce.columns:
            df_cruce['Precio'] = df_cruce['MATERIAL'].apply(lambda x: len(str(x)) * 1250.50) # Precio inventado
            simulado = True
        if 'Venta_Mensual' not in df_cruce.columns:
            df_cruce['Venta_Mensual'] = df_cruce['MATERIAL'].apply(lambda x: len(str(x)) * 300) # Venta inventada
            simulado = True
            
        if simulado:
            st.warning("⚠️ **Modo Demo:** Como no encontré las columnas 'Precio' y 'Venta_Mensual' en SAP, el sistema generó valores ficticios para mostrarte cómo se ve. Para usar datos reales, simplemente agrega esas dos columnas a tu archivo SAP.")

        # MOTOR 3: Valorización Económica del Riesgo
        df_cruce['Valor_Stock'] = df_cruce['Total'] * df_cruce['Precio']
        
        # MOTOR 4: Días de Cobertura (Stock Cover)
        df_cruce['Venta_Diaria'] = df_cruce['Venta_Mensual'] / 30
        df_cruce['Cobertura_Dias'] = df_cruce.apply(lambda r: round(r['Total'] / r['Venta_Diaria']) if r['Venta_Diaria'] > 0 else 0, axis=1)

        # ==========================================
        # CREACIÓN DE PESTAÑAS GERENCIALES
        # ==========================================
        tab1, tab2, tab3, tab4, tab5 = st.tabs([
            "📊 Resumen", "🧊 Ocupación (Pallets)", "🚨 Alertas FEFO y Bloqueos", "💰 Riesgo y Cobertura", "📋 Base de Datos"
        ])

        # --- TAB 1: RESUMEN ---
        with tab1:
            st.markdown("### 📈 Indicadores Globales")
            k1, k2, k3, k4 = st.columns(4)
            k1.metric("📦 Unidades Totales", f"{df_cruce['Total'].sum():,.0f}")
            k2.metric("🧊 Pallets Físicos", f"{df_cruce['Pallets_Ocupados'].sum():,.1f}")
            k3.metric("🔴 Lotes Críticos", len(df_cruce[df_cruce['Estado_Alerta'] == '2. Crítico (<30 días)']))
            riesgo_dinero = df_cruce[df_cruce['Estado_Alerta'].str.contains('1|2', na=False)]['Valor_Stock'].sum()
            k4.metric("⚠️ Dinero en Riesgo (<30 días)", f"${riesgo_dinero:,.2f}")

            df_pie = df_cruce[df_cruce['Estado_Alerta'] != '5. Sin Fecha']
            mapa_colores = {"1. Vencido": "#8B0000", "2. Crítico (<30 días)": "#FF0000", "3. Alerta (31-60 días)": "#FFA500", "4. Precaución (61-90 días)": "#FFD700", "0. Óptimo (>90 días)": "#228B22"}
            if not df_pie.empty:
                fig_pie = px.pie(df_pie, names='Estado_Alerta', color='Estado_Alerta', color_discrete_map=mapa_colores, hole=0.4, title="Distribución de Vencimientos")
                st.plotly_chart(fig_pie, use_container_width=True)

        # --- TAB 2: OCUPACIÓN FÍSICA ---
        with tab2:
            st.markdown("### 🧊 Análisis de Capacidad Instalada")
            st.info("Aquí identificamos qué materiales están saturando las posiciones del almacén físico.")
            top_pallets = df_cruce.groupby('NOMBRE_x')['Pallets_Ocupados'].sum().reset_index().sort_values(by='Pallets_Ocupados', ascending=False).head(15)
            fig_pallets = px.bar(top_pallets, x='Pallets_Ocupados', y='NOMBRE_x', orientation='h', title="Top 15 Productos que más Pallets Ocupan")
            fig_pallets.update_layout(yaxis={'categoryorder':'total ascending'})
            st.plotly_chart(fig_pallets, use_container_width=True)

        # --- TAB 3: ALERTAS FEFO Y BLOQUEOS ---
        with tab3:
            st.markdown("### 🛑 Centro de Control Operativo")
            col_a, col_b = st.columns(2)
            with col_a:
                st.subheader("1. Bloqueo Rápido SAP")
                umbral = st.slider("Ver stock menor a:", 0, 500, 50, 10)
                df_bloqueo = df_cruce[df_cruce['Total'] <= umbral][['MATERIAL', 'NOMBRE_x', 'Total', 'Estado_Alerta']]
                st.dataframe(df_bloqueo.sort_values(by='Total'), hide_index=True)
            
            with col_b:
                st.subheader("2. Alerta FEFO (Múltiples Lotes)")
                st.markdown("Materiales con stock activo en **más de un lote diferente**. Riesgo de mal despacho.")
                df_activos = df_cruce[df_cruce['Total'] > 0]
                conteo_lotes = df_activos.groupby(['MATERIAL', 'NOMBRE_x'])['LOTE'].nunique().reset_index()
                fefo_alert = conteo_lotes[conteo_lotes['LOTE'] > 1].rename(columns={'LOTE': 'Cant. Lotes Distintos'})
                st.dataframe(fefo_alert, hide_index=True)

        # --- TAB 4: RIESGO Y COBERTURA ---
        with tab4:
            st.markdown("### 💰 Impacto Financiero y Ventas")
            st.markdown("Cruza el almacén con el ritmo de ventas para detectar sobrestock o quiebres inminentes.")
            
            riesgo_df = df_cruce[['NOMBRE_x', 'Total', 'Cobertura_Dias', 'Valor_Stock', 'Estado_Alerta']].sort_values(by='Valor_Stock', ascending=False).head(20)
            riesgo_df['Valor_Stock'] = riesgo_df['Valor_Stock'].apply(lambda x: f"${x:,.2f}")
            st.dataframe(riesgo_df, hide_index=True, use_container_width=True)

        # --- TAB 5: DESCARGA ---
        with tab5:
            st.markdown("### 📋 Tabla Maestra")
            cols_final = ['MATERIAL', 'NOMBRE_x', 'LOTE', 'Fecha_Mostrada', 'Dias_Para_Vencer', 'Total', 'Pallets_Ocupados', 'Cobertura_Dias', 'Valor_Stock', 'Estado_Alerta']
            df_export = df_cruce[cols_final].copy()
            df_export.rename(columns={'NOMBRE_x': 'Descripción'}, inplace=True)
            st.dataframe(df_export.head(100), hide_index=True) # Mostrar solo 100 para no trabar el navegador
            
            def generar_excel_completo(df):
                wb = Workbook()
                ws = wb.active
                ws.title = "Maestro Gerencial"
                for r in dataframe_to_rows(df, index=False, header=True):
                    ws.append(r)
                
                # Semaforización en Excel
                fill_rojo = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
                fill_nar = PatternFill(start_color="FFE6CC", end_color="FFE6CC", fill_type="solid")
                fill_ama = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
                
                idx_estado = len(df.columns)
                for row in ws.iter_rows(min_row=2):
                    estado_val = str(row[idx_estado-1].value)
                    for cell in row:
                        if "1." in estado_val or "2." in estado_val: cell.fill = fill_rojo
                        elif "3." in estado_val: cell.fill = fill_nar
                        elif "4." in estado_val: cell.fill = fill_ama
                
                output = io.BytesIO()
                wb.save(output)
                return output.getvalue()

            st.download_button("📥 Descargar Excel Estratégico", data=generar_excel_completo(df_export), file_name=f"Estrategico_La_Virginia_{datetime.date.today()}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    except Exception as e:
        st.error(f"Error procesando datos: {e}")
