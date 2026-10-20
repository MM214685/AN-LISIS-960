# -*- coding: utf-8 -*-
"""
Created on Mon Oct 19 09:47:39 2026

@author: meseg
"""

import streamlit as st
import pandas as pd
import numpy as np
import io
import plotly.express as px

# --- Configuración de la Web ---
st.set_page_config(page_title="Dashboard Rentabilidad Antigua GEMA", layout="wide")
st.title("📊 Cuadro de Mandos: Rentabilidad Antigua GEMA")

st.markdown("""
Sube tu archivo de datos para visualizar el rendimiento y generar el **Excel automatizado** con todas las reglas de negocio.
""")

# --- Subida de Archivos ---
uploaded_file = st.file_uploader("Sube el archivo de datos (Excel/CSV)", type=['xlsx', 'xls', 'csv'])

if uploaded_file:
    # 1. LECTURA DE DATOS
    try:
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file, skiprows=16)
        else:
            df = pd.read_excel(uploaded_file, skiprows=16)
            
        if 'Ventas' not in df.columns and 'Linea de Venta' not in df.columns:
            if uploaded_file.name.endswith('.csv'):
                df = pd.read_csv(uploaded_file)
            else:
                df = pd.read_excel(uploaded_file)
                
    except Exception as e:
        st.error(f"Error al leer el archivo: {e}")
        st.stop()

    # --- LIMPIEZA Y PREPARACIÓN DE DATOS ---
    try:
        # Separar el Año y el Mes
        if 'Mes' in df.columns and df['Mes'].astype(str).str.contains('-').any():
            df[['Año', 'Mes_Solo']] = df['Mes'].str.split('-', expand=True)
            df['Año'] = pd.to_numeric(df['Año'], errors='coerce')
            df['Mes'] = df['Mes_Solo']
            
        # Renombrar columnas
        rename_dict = {
            'Linea de Venta': 'Linea_Venta',
            'Nombre Proyecto': 'Contrato', 
            'Oficina': 'Provincia',
            'Coste Directo': 'Costes'
        }
        df = df.rename(columns=rename_dict)

        # UNIFICACIÓN DE CONTRATOS UBE
        def unificar_ube(nombre):
            if not isinstance(nombre, str): return nombre
            n_low = nombre.lower()
            if 'ube' in n_low and ('nuevo cif' in n_low or 'almacen' in n_low or 'almacén' in n_low):
                return 'UBE (Consolidado almacenes/nuevo CIF)'
            return nombre
            
        df['Contrato'] = df['Contrato'].apply(unificar_ube)

        # Convertir a números
        df['Ventas'] = pd.to_numeric(df['Ventas'], errors='coerce').fillna(0)
        df['Costes'] = pd.to_numeric(df['Costes'], errors='coerce').fillna(0)
        
        # Calcular márgenes globales
        df['Margen €'] = df['Ventas'] - df['Costes']
        df['Margen %'] = np.where(df['Ventas'] != 0, df['Margen €'] / df['Ventas'], 0)
        
        # Categorizar meses cronológicamente
        meses_orden = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre']
        df['Mes'] = pd.Categorical(df['Mes'], categories=meses_orden, ordered=True)
        mes2num = {m: i+1 for i, m in enumerate(meses_orden)}
        df['Mes_Num'] = df['Mes'].map(mes2num)

    except KeyError as e:
        st.error(f"⚠️ No se encontró la columna esperada. Detalles: {e}")
        st.stop()

    # --- FILTROS INTERACTIVOS ---
    st.markdown("---")
    st.markdown("### 🎛️ Filtros de Análisis")
    col_f1, col_f2 = st.columns(2)
    with col_f1:
        temporada_sel = st.selectbox("🗓️ Temporada a visualizar", ["Ambas (Comparativa)", 2026, 2025])
    with col_f2:
        mes_corte = st.selectbox("📅 Mes de corte (Acumulado hasta...)", meses_orden, index=8) # Por defecto Septiembre (9)
    st.markdown("---")

    # Aplicar filtros
    df_filtrado = df[df['Mes_Num'] <= mes2num[mes_corte]]
    
    df_2025 = df_filtrado[df_filtrado['Año'] == 2025]
    df_2026 = df_filtrado[df_filtrado['Año'] == 2026]
    
    mostrar_25 = temporada_sel in ["Ambas (Comparativa)", 2025]
    mostrar_26 = temporada_sel in ["Ambas (Comparativa)", 2026]

    # --- FUNCIONES DE ESTILO Y CÁLCULO ---
    def style_df(df_to_style, is_multi=False):
        if df_to_style.empty: return df_to_style
        f_dict = {}
        if is_multi:
            for col in df_to_style.columns:
                if 'Margen %' in col[1]: f_dict[col] = "{:.2%}"
                else: f_dict[col] = "{:,.2f} €"
        else:
            for col in df_to_style.columns:
                if 'Margen %' in col: f_dict[col] = "{:.2%}"
                elif col in ['Ventas', 'Costes', 'Margen €']: f_dict[col] = "{:,.2f} €"
                
        def color_negative(val):
            if isinstance(val, (int, float)) and val < 0: return 'color: red; font-weight: bold;'
            return ''
            
        def highlight_total(row):
            if row.name == 'TOTAL' or (isinstance(row.name, str) and row.name == 'TOTAL'):
                return ['background-color: #f0f2f6; font-weight: bold'] * len(row)
            if 'Linea_Venta' in row and row['Linea_Venta'] == 'TOTAL':
                return ['background-color: #f0f2f6; font-weight: bold'] * len(row)
            return [''] * len(row)

        try:
            return df_to_style.style.format(f_dict).applymap(color_negative).apply(highlight_total, axis=1)
        except AttributeError:
            return df_to_style.style.format(f_dict).map(color_negative).apply(highlight_total, axis=1)

    def generar_tabla_mensual(df_year):
        if df_year.empty: return pd.DataFrame()
        agrupado = df_year.groupby(['Linea_Venta', 'Mes']).agg({'Ventas':'sum', 'Costes':'sum'}).reset_index()
        agrupado = agrupado[(agrupado['Ventas'] != 0) | (agrupado['Costes'] != 0)]
        
        agrupado['Margen €'] = agrupado['Ventas'] - agrupado['Costes']
        agrupado['Margen %'] = np.where(agrupado['Ventas']!=0, agrupado['Margen €']/agrupado['Ventas'], 0)
        
        pivot = agrupado.pivot(index='Linea_Venta', columns='Mes', values=['Ventas', 'Costes', 'Margen €', 'Margen %'])
        pivot = pivot.dropna(axis=1, how='all')
        pivot = pivot.swaplevel(0, 1, axis=1).sort_index(axis=1, level=0)
        
        metric_order = ['Ventas', 'Costes', 'Margen €', 'Margen %']
        pivot = pivot.reindex(metric_order, level=1, axis=1)
        
        # Añadir fila de totales
        meses_presentes = pivot.columns.get_level_values(0).unique()
        for mes in meses_presentes:
            v = pivot[(mes, 'Ventas')].sum(skipna=True)
            c = pivot[(mes, 'Costes')].sum(skipna=True)
            m_eur = v - c
            m_pct = m_eur / v if v != 0 else 0
            pivot.loc['TOTAL', (mes, 'Ventas')] = v
            pivot.loc['TOTAL', (mes, 'Costes')] = c
            pivot.loc['TOTAL', (mes, 'Margen €')] = m_eur
            pivot.loc['TOTAL', (mes, 'Margen %')] = m_pct
            
        return pivot

    def calcular_top_bottom(df_year):
        if df_year.empty: return pd.DataFrame(), pd.DataFrame()
        agrupado = df_year.groupby(['Contrato', 'Linea_Venta']).agg({'Ventas':'sum', 'Costes':'sum'}).reset_index()
        agrupado = agrupado[(agrupado['Ventas'] != 0) | (agrupado['Costes'] != 0)]
        agrupado['Margen €'] = agrupado['Ventas'] - agrupado['Costes']
        agrupado['Margen %'] = np.where(agrupado['Ventas']!=0, agrupado['Margen €']/agrupado['Ventas'], 0)
        
        agrupado = agrupado.sort_values(by=['Margen €', 'Ventas', 'Margen %'], ascending=[False, False, False])
        
        top_10 = agrupado.head(10).reset_index(drop=True)
        top_10.index += 1
        
        bottom_10 = agrupado.tail(10).sort_values(by=['Margen €', 'Ventas', 'Margen %'], ascending=[True, True, True]).reset_index(drop=True)
        bottom_10.index += 1
        
        return top_10, bottom_10

    def calcular_acumulado(df_year):
        if df_year.empty: return pd.DataFrame()
        agrupado = df_year.groupby('Linea_Venta').agg({'Ventas':'sum', 'Costes':'sum'}).reset_index()
        agrupado['Margen €'] = agrupado['Ventas'] - agrupado['Costes']
        agrupado['Margen %'] = np.where(agrupado['Ventas']!=0, agrupado['Margen €']/agrupado['Ventas'], 0)
        
        tot_v = agrupado['Ventas'].sum()
        tot_c = agrupado['Costes'].sum()
        tot_me = tot_v - tot_c
        tot_mp = tot_me / tot_v if tot_v != 0 else 0
        agrupado.loc[len(agrupado)] = ['TOTAL', tot_v, tot_c, tot_me, tot_mp]
        
        return agrupado

    # --- PESTAÑAS DE LA WEB ---
    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "Resultados Mensuales", f"Acumulado (hasta {mes_corte})", "Top/Bottom Contratos", "Resumen Provincias", "📥 Generar y Descargar Excel"
    ])

    with tab1:
        st.subheader(f"Tablas Mensuales por Línea de Venta (Hasta {mes_corte})")
        
        if mostrar_25:
            st.markdown("### 📅 Temporada 2025")
            df_m25 = generar_tabla_mensual(df_2025)
            if not df_m25.empty:
                # Calculamos la altura perfecta: nº de filas * 36px + cabeceras + padding
                altura_25 = len(df_m25) * 38 + 90
                st.dataframe(style_df(df_m25, is_multi=True), use_container_width=True, height=altura_25)
                
        if mostrar_26:
            st.markdown("---")
            st.markdown("### 📅 Temporada 2026")
            df_m26 = generar_tabla_mensual(df_2026)
            if not df_m26.empty:
                altura_26 = len(df_m26) * 38 + 90
                st.dataframe(style_df(df_m26, is_multi=True), use_container_width=True, height=altura_26)

    with tab2:
        st.subheader(f"Comparativa Acumulada por Línea de Venta (Enero - {mes_corte})")
        if mostrar_25:
            st.write("2025:")
            df_a25 = calcular_acumulado(df_2025)
            if not df_a25.empty:
                altura_a25 = len(df_a25) * 38 + 50
                st.dataframe(style_df(df_a25), use_container_width=True, height=altura_a25)
                
        if mostrar_26:
            st.write("2026:")
            df_a26 = calcular_acumulado(df_2026)
            if not df_a26.empty:
                altura_a26 = len(df_a26) * 38 + 50
                st.dataframe(style_df(df_a26), use_container_width=True, height=altura_a26)

    with tab3:
        st.subheader(f"10 Mejores y 10 Peores Contratos (Acumulado hasta {mes_corte})")
        top_25, bot_25 = calcular_top_bottom(df_2025)
        top_26, bot_26 = calcular_top_bottom(df_2026)
        
        if mostrar_26:
            st.markdown("### 📅 Temporada 2026")
            c1, c2 = st.columns(2)
            with c1:
                st.markdown("🥇 **Top 10 - 2026**")
                st.dataframe(style_df(top_26), use_container_width=True, height=430)
            with c2:
                st.markdown("📉 **Peores 10 - 2026**")
                st.dataframe(style_df(bot_26), use_container_width=True, height=430)
            
        if mostrar_25:
            st.markdown("---")
            st.markdown("### 📅 Temporada 2025")
            c3, c4 = st.columns(2)
            with c3:
                st.markdown("🥇 **Top 10 - 2025**")
                st.dataframe(style_df(top_25), use_container_width=True, height=430)
            with c4:
                st.markdown("📉 **Peores 10 - 2025**")
                st.dataframe(style_df(bot_25), use_container_width=True, height=430)

    with tab4:
        st.subheader(f"Resumen por Provincias (Hasta {mes_corte})")
        df_prov_filtrado = df_filtrado.copy()
        if temporada_sel != "Ambas (Comparativa)":
            df_prov_filtrado = df_prov_filtrado[df_prov_filtrado['Año'] == temporada_sel]
            
        if not df_prov_filtrado.empty:
            res_prov = df_prov_filtrado.groupby(['Año', 'Provincia']).agg({'Ventas':'sum', 'Costes':'sum'}).reset_index()
            res_prov['Margen €'] = res_prov['Ventas'] - res_prov['Costes']
            res_prov['Margen %'] = np.where(res_prov['Ventas']!=0, res_prov['Margen €']/res_prov['Ventas'], 0)
            
            fig = px.bar(res_prov, x="Provincia", y="Margen €", color="Año", barmode="group", title=f"Margen Absoluto por Provincia (Acumulado a {mes_corte})")
            st.plotly_chart(fig, use_container_width=True)
            st.dataframe(style_df(res_prov), use_container_width=True)
        else:
            st.write("No hay datos para esta selección.")

    with tab5:
        st.subheader("Generación de Archivo Excel Final")
        st.markdown(f"*El Excel se generará aplicando el filtro seleccionado: **Acumulado hasta {mes_corte}**.*")
        st.markdown("*Pega aquí las conclusiones que te genere la IA en nuestra conversación para integrarlas al Excel.*")
        conclusiones_texto = st.text_area("Conclusiones y Análisis de Tendencia:", height=150)
        
        if st.button("Generar Archivo Excel"):
            output = io.BytesIO()
            writer = pd.ExcelWriter(output, engine='xlsxwriter')
            workbook = writer.book
            
            fmt_moneda = workbook.add_format({'num_format': '#,##0.00 €;[Red]-#,##0.00 €'})
            fmt_pct = workbook.add_format({'num_format': '0.00%;[Red]-0.00%'})
            fmt_cabecera = workbook.add_format({'bold': True, 'bg_color': '#4F81BD', 'font_color': 'white', 'align': 'center', 'valign': 'vcenter'})
            fmt_total = workbook.add_format({'bold': True, 'bg_color': '#D9D9D9'})

            # 1. Hoja: Resultados Mensuales
            df_mensual_25 = generar_tabla_mensual(df_2025)
            df_mensual_26 = generar_tabla_mensual(df_2026)
            
            fila_inicio_26 = 1
            if mostrar_25 and not df_mensual_25.empty:
                df_mensual_25.to_excel(writer, sheet_name='Resultados Mensuales', startrow=1)
                ws_mensual = writer.sheets['Resultados Mensuales']
                ws_mensual.write_string(0, 0, "RESULTADOS 2025", fmt_cabecera)
                
                fila_total_25 = len(df_mensual_25) + 2
                ws_mensual.set_row(fila_total_25, None, fmt_total)
                
                fila_inicio_26 = len(df_mensual_25) + 5
                
            if mostrar_26 and not df_mensual_26.empty:
                df_mensual_26.to_excel(writer, sheet_name='Resultados Mensuales', startrow=fila_inicio_26)
                ws_mensual = writer.sheets['Resultados Mensuales']
                ws_mensual.write_string(fila_inicio_26 - 1, 0, "RESULTADOS 2026", fmt_cabecera)
                
                fila_total_26 = fila_inicio_26 + len(df_mensual_26) + 1
                ws_mensual.set_row(fila_total_26, None, fmt_total)
            
            if 'Resultados Mensuales' in writer.sheets:
                ws_mensual = writer.sheets['Resultados Mensuales']
                ws_mensual.set_column(0, 0, 45)
                for col_num in range(1, 50):
                    if col_num % 4 == 0: ws_mensual.set_column(col_num, col_num, 12, fmt_pct)
                    else: ws_mensual.set_column(col_num, col_num, 15, fmt_moneda)
            
            # 2. Hoja: Acumulado
            df_acum_25 = calcular_acumulado(df_2025)
            df_acum_26 = calcular_acumulado(df_2026)
            
            fila_inicio_26 = 1
            if mostrar_25 and not df_acum_25.empty:
                df_acum_25.to_excel(writer, sheet_name='Acumulado', index=False, startrow=1)
                ws_acum = writer.sheets['Acumulado']
                ws_acum.write_string(0, 0, "ACUMULADO 2025", fmt_cabecera)
                
                fila_total_25 = len(df_acum_25) + 1
                ws_acum.set_row(fila_total_25, None, fmt_total)
                
                fila_inicio_26 = len(df_acum_25) + 5
                
            if mostrar_26 and not df_acum_26.empty:
                df_acum_26.to_excel(writer, sheet_name='Acumulado', index=False, startrow=fila_inicio_26)
                ws_acum = writer.sheets['Acumulado']
                ws_acum.write_string(fila_inicio_26 - 1, 0, "ACUMULADO 2026", fmt_cabecera)
                
                fila_total_26 = fila_inicio_26 + len(df_acum_26)
                ws_acum.set_row(fila_total_26, None, fmt_total)
                
            if 'Acumulado' in writer.sheets:
                ws_acum = writer.sheets['Acumulado']
                ws_acum.set_column(0, 0, 45)
                ws_acum.set_column(1, 3, 16, fmt_moneda)
                ws_acum.set_column(4, 4, 12, fmt_pct)
            
            # 3. Hoja: Contratos
            fila_actual = 1
            if mostrar_25:
                top_25.to_excel(writer, sheet_name='Top Bottom Contratos', index=True, startrow=fila_actual)
                ws_tb = writer.sheets['Top Bottom Contratos']
                ws_tb.write_string(fila_actual-1, 0, "🥇 Top 10 - 2025", fmt_cabecera)
                fila_actual += 15
                
                bot_25.to_excel(writer, sheet_name='Top Bottom Contratos', index=True, startrow=fila_actual)
                ws_tb.write_string(fila_actual-1, 0, "📉 Peores 10 - 2025", fmt_cabecera)
                fila_actual += 15
                
            if mostrar_26:
                top_26.to_excel(writer, sheet_name='Top Bottom Contratos', index=True, startrow=fila_actual)
                ws_tb = writer.sheets['Top Bottom Contratos']
                ws_tb.write_string(fila_actual-1, 0, "🥇 Top 10 - 2026", fmt_cabecera)
                fila_actual += 15
                
                bot_26.to_excel(writer, sheet_name='Top Bottom Contratos', index=True, startrow=fila_actual)
                ws_tb.write_string(fila_actual-1, 0, "📉 Peores 10 - 2026", fmt_cabecera)

            if 'Top Bottom Contratos' in writer.sheets:
                ws_tb = writer.sheets['Top Bottom Contratos']
                ws_tb.set_column(0, 0, 5) 
                ws_tb.set_column(1, 1, 50) 
                ws_tb.set_column(2, 2, 40) 
                ws_tb.set_column(3, 5, 16, fmt_moneda)
                ws_tb.set_column(6, 6, 12, fmt_pct)
            
            # 4. Hoja: Provincias
            if not df_prov_filtrado.empty:
                res_prov.to_excel(writer, sheet_name='Provincias', index=False)
                ws_prov = writer.sheets['Provincias']
                ws_prov.set_column(0, 1, 15)
                ws_prov.set_column(2, 4, 16, fmt_moneda)
                ws_prov.set_column(5, 5, 12, fmt_pct)
            
            # 5. Hoja: Conclusiones
            ws_conc = workbook.add_worksheet("Conclusiones")
            ws_conc.write(0, 0, f"Conclusiones del Cierre Mensual (Hasta {mes_corte}) y Tendencias", fmt_cabecera)
            ws_conc.write(2, 0, conclusiones_texto)
            ws_conc.set_column(0, 0, 150) 

            writer.close()
            output.seek(0)
            
            st.success(f"¡Archivo generado con éxito! (Datos acumulados hasta {mes_corte})")
            st.download_button(
                label="📥 Descargar Dashboard (Excel)",
                data=output,
                file_name=f"Dashboard_Rentabilidad_GEMA_Hasta_{mes_corte}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )