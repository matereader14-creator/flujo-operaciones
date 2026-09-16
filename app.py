import os
import io
import streamlit as st
import pandas as pd
import plotly.express as px

# 1. Configuración principal de la página
st.set_page_config(page_title="Flujo de Operaciones", layout="wide", page_icon="📊")

# Título con el logo de Toyota ampliado
st.markdown(
    """
    <div style="display: flex; align-items: center; margin-bottom: 20px;">
        <img src="https://upload.wikimedia.org/wikipedia/commons/e/e7/Toyota.svg" width="130" style="margin-right: 20px;">
        <h1 style="margin: 0;">Flujo y Seguimiento de Operaciones</h1>
    </div>
    """, 
    unsafe_allow_html=True
)

# ==========================================
# 2. CONEXIÓN A GOOGLE SHEETS EN LA NUBE
# ==========================================
URL_GOOGLE_SHEETS = "https://docs.google.com/spreadsheets/d/1gvdyoiorFXTNiXREPe2Xt55mOEjmUkAXwUEfOCmecfA/edit?usp=sharing"

@st.cache_data(ttl=600) # Se actualiza cada 10 minutos automáticamente
def cargar_datos_desde_sheets(url):
    try:
        if "edit" in url:
            url_descarga = url.split("/edit")[0] + "/export?format=xlsx"
        else:
            url_descarga = url
            
        return pd.read_excel(url_descarga, sheet_name='Boleto')
    except Exception as e:
        st.error(f"Error al conectar con Google Sheets: {e}")
        return None

df_cargado = None

if "docs.google.com" in URL_GOOGLE_SHEETS:
    with st.spinner("Sincronizando base de datos en vivo..."):
        df_cargado = cargar_datos_desde_sheets(URL_GOOGLE_SHEETS)
else:
    st.sidebar.warning("Link de Google Sheets no configurado. Sube el archivo manualmente.")
    archivo_subido = st.sidebar.file_uploader("Sube tu base de datos (Excel):", type=["xlsx", "xls"])
    if archivo_subido is not None:
        df_cargado = pd.read_excel(archivo_subido, sheet_name='Boleto')

# ==========================================
# 3. PROCESAMIENTO Y GRÁFICOS
# ==========================================
if df_cargado is not None:
    try:
        df = df_cargado.copy()
        
        # --- CREACIÓN DE IDENTIFICADOR ÚNICO ACTUALIZADO ---
        if 'Nombres_de_Cuenta__c' in df.columns and 'Denominacion_Comercial__c' in df.columns:
            nombres = df['Nombres_de_Cuenta__c'].fillna("Cliente Sin Nombre")
            vehiculos = df['Denominacion_Comercial__c'].fillna("Vehículo Sin Asignar")
            
            if 'Numero_de_Boleto__c' in df.columns:
                boletos = df['Numero_de_Boleto__c'].fillna("S/N").astype(str)
                df['Identificador'] = nombres.astype(str) + " | " + vehiculos.astype(str) + " | Boleto: " + boletos
            else:
                df['Identificador'] = nombres.astype(str) + " | " + vehiculos.astype(str)
        else:
            df['Identificador'] = df.index.astype(str)
            
        # --- CARGA DE COMENTARIOS PERMANENTES ---
        ARCHIVO_COMENTARIOS = 'comentarios_operaciones.csv'
        if os.path.exists(ARCHIVO_COMENTARIOS):
            df_com = pd.read_csv(ARCHIVO_COMENTARIOS)
            df = df.merge(df_com, on='Identificador', how='left')
        else:
            df['Comentario'] = ""
            
        df['Comentario'] = df['Comentario'].fillna("")
        
        # Procesamiento de Columnas de Fecha
        cols_fechas = [c for c in df.columns if 'fecha' in str(c).lower()]
        
        for col in cols_fechas:
            df[col] = pd.to_datetime(df[col], errors='coerce')
            
        # FILTRO AÑO 2026 
        mask_2026 = pd.Series(False, index=df.index)
        for col in cols_fechas:
            mask_2026 = mask_2026 | (df[col].dt.year == 2026)
        
        df = df[mask_2026].copy()
        
        # Diccionario de tiempos límite de demora
        limites_demora = {
            "Ingreso Pendiente": 1,
            "Creacion En Siac": 1,
            "Proboleto Aprobado": 2,
            "Pedido Confirmado": 5,
            "Firma Del Boleto": 5,
            "En Proceso De Pago": 5,
            "Facturacion": 3,
            "Poliza De Seguro": 2,
            "En Proceso De Patentamiento": 15,
            "En Proceso De Entrega": 1
        }
        
        # Lógica de Negocio: Estado actual y demoras
        def calcular_estado_y_demora(row):
            estado_actual = "Ingreso Pendiente"
            ultima_fecha = pd.NaT
            
            columnas_ordenadas = [
                'Fecha_de_creacion_en_SIAC__c', 'Fecha_de_proboleto_aprobado__c', 'Fecha_de_pedido_confirmado__c', 
                'Fecha_de_firma_del_boleto__c', 'Fecha_de_en_proceso_de_pago__c', 'Fecha_de_Facturacion__c', 
                'Fecha_de_Poliza_de_Seguro__c', 'Fecha_en_Proceso_de_Patentamiento__c', 
                'Fecha_de_en_proceso_de_entrega__c', 
                'Fecha_de_disfruta_tu_nuevo_Toyota__c'
            ]
            
            for col in columnas_ordenadas:
                if col in df.columns and pd.notna(row[col]):
                    estado_limpio = str(col).replace('Fecha_de_', '').replace('Fecha_', '').replace('__c', '').replace('_', ' ').title()
                    estado_actual = estado_limpio
                    ultima_fecha = row[col]
            
            es_baja = False
            
            if 'Fecha_de_baja__c' in df.columns and pd.notna(row['Fecha_de_baja__c']):
                estado_actual = "Operación Dada De Baja"
                ultima_fecha = row['Fecha_de_baja__c']
                es_baja = True
                
            elif 'Denominacion_Comercial__c' in df.columns:
                vehiculo = str(row['Denominacion_Comercial__c']).strip().lower()
                if pd.isna(row['Denominacion_Comercial__c']) or vehiculo in ('nan', 'none', ''):
                    estado_actual = "Operación Dada De Baja"
                    es_baja = True
                    
            if not es_baja:
                tiene_pro = 'Fecha_de_proboleto_aprobado__c' in df.columns and pd.notna(row['Fecha_de_proboleto_aprobado__c'])
                tiene_bol = 'Fecha_de_firma_del_boleto__c' in df.columns and pd.notna(row['Fecha_de_firma_del_boleto__c'])
                
                if not tiene_pro and not tiene_bol:
                    tiene_progreso_real = False
                    for col in cols_fechas:
                        nombre_col_min = str(col).lower()
                        if 'entrega' not in nombre_col_min and 'disfruta' not in nombre_col_min and 'baja' not in nombre_col_min and 'creacion' not in nombre_col_min:
                            if pd.notna(row[col]):
                                tiene_progreso_real = True
                                break
                    if not tiene_progreso_real:
                        estado_actual = "Operación Dada De Baja"
                        es_baja = True

            if not es_baja:
                if estado_actual == "Proboleto Aprobado" and pd.notna(ultima_fecha):
                    if (pd.Timestamp.now() - ultima_fecha).days > 5:
                        estado_actual = "Operación Dada De Baja"
                        es_baja = True
            
            if pd.notna(ultima_fecha):
                if estado_actual == "Disfruta Tu Nuevo Toyota" or es_baja:
                    demora = 0 
                else:
                    demora = (pd.Timestamp.now() - ultima_fecha).days
            else:
                demora = 0
                
            return pd.Series([estado_actual, demora, ultima_fecha])
            
        # --- PARACAÍDAS PARA TABLAS VACÍAS ---
        if df.empty:
            df['Estado Actual'] = pd.Series(dtype='object')
            df['Días en este estado'] = pd.Series(dtype='float64')
            df['Fecha de Último Estado'] = pd.Series(dtype='datetime64[ns]')
            df['Mes'] = pd.Series(dtype='object')
            st.warning("⚠️ No se encontraron operaciones del año 2026 en la base de datos actual.")
        else:
            df[['Estado Actual', 'Días en este estado', 'Fecha de Último Estado']] = df.apply(calcular_estado_y_demora, axis=1)
            df['Mes'] = df['Fecha de Último Estado'].dt.to_period('M').astype(str)
        
        # MÁSCARA GLOBAL DE RETRASOS
        if df.empty:
            retrasados_mask = pd.Series(False, index=df.index)
        else:
            retrasados_mask = df.apply(
                lambda x: x['Días en este estado'] > limites_demora.get(x['Estado Actual'], 5) 
                if x['Estado Actual'] not in ['Disfruta Tu Nuevo Toyota', 'Operación Dada De Baja'] else False, 
                axis=1
            )
        
        # FILTROS EN LA BARRA LATERAL
        st.sidebar.markdown("---")
        st.sidebar.subheader("Filtros Globales")
        
        meses_limpios = [str(m) for m in df['Mes'].unique() if str(m) not in ('NaT', 'nan', 'None')]
        meses_disponibles = sorted(meses_limpios)
        opciones_mes = ["Todos los meses"] + meses_disponibles
        
        mes_actual = pd.Timestamp.now().strftime('%Y-%m') 
        idx_mes_defecto = 0
        if mes_actual in opciones_mes:
            idx_mes_defecto = opciones_mes.index(mes_actual)
            
        mes_seleccionado = st.sidebar.selectbox("Seleccionar Mes (2026):", opciones_mes, index=idx_mes_defecto)
        
        if mes_seleccionado != "Todos los meses":
            df = df[df['Mes'] == mes_seleccionado].copy()
            retrasados_mask = df.apply(
                lambda x: x['Días en este estado'] > limites_demora.get(x['Estado Actual'], 5) 
                if x['Estado Actual'] not in ['Disfruta Tu Nuevo Toyota', 'Operación Dada De Baja'] else False, 
                axis=1
            )

        if 'Sucursal_de_Venta__c' in df.columns:
            sucursales_limpias = [str(s) for s in df['Sucursal_de_Venta__c'].unique() if str(s) not in ('nan', 'None', 'NaT') and pd.notna(s)]
            sucursales_disponibles = sorted(sucursales_limpias)
            
            if sucursales_disponibles:
                opciones_sucursal = ["Todas las sucursales"] + sucursales_disponibles
                idx_sucursal_defecto = 0
                for i, suc in enumerate(opciones_sucursal):
                    if "autolux salta" in suc.lower():
                        idx_sucursal_defecto = i
                        break
                        
                sucursal_seleccionada = st.sidebar.selectbox("Seleccionar Sucursal:", opciones_sucursal, index=idx_sucursal_defecto)
                
                if sucursal_seleccionada != "Todas las sucursales":
                    df = df[df['Sucursal_de_Venta__c'].astype(str) == sucursal_seleccionada].copy()
                    retrasados_mask = df.apply(
                        lambda x: x['Días en este estado'] > limites_demora.get(x['Estado Actual'], 5) 
                        if x['Estado Actual'] not in ['Disfruta Tu Nuevo Toyota', 'Operación Dada De Baja'] else False, 
                        axis=1
                    )
        
        def convertir_df_a_excel(df_export):
            salida = io.BytesIO()
            with pd.ExcelWriter(salida, engine='openpyxl') as writer:
                df_export.to_excel(writer, index=False, sheet_name='Datos_Boletos')
            return salida.getvalue()

        # Creación de Pestañas Visuales
        tab1, tab2, tab3, tab4, tab5 = st.tabs([
            "👥 Rendimiento del Equipo", 
            "🔍 Auditoría y Detalles",
            "🔎 Rastreo Individual",
            "📊 Estado de Operaciones",
            "⚠️ Alertas de Demora"
        ])
        
        with tab1:
            st.subheader("Carga de Trabajo Operativo")
            colA, colB = st.columns(2)
            with colA:
                if 'Nombre_Asesor__c' in df.columns and not df.empty:
                    vendedores = df.groupby('Nombre_Asesor__c').size().reset_index(name='Operaciones')
                    fig_vendedores = px.pie(vendedores, names='Nombre_Asesor__c', values='Operaciones', title='Operaciones por Asesor', hole=0.4)
                    st.plotly_chart(fig_vendedores, use_container_width=True)
                elif 'Asesor__c' in df.columns and not df.empty: 
                    vendedores = df.groupby('Asesor__c').size().reset_index(name='Operaciones')
                    fig_vendedores = px.pie(vendedores, names='Asesor__c', values='Operaciones', title='Operaciones por Asesor', hole=0.4)
                    st.plotly_chart(fig_vendedores, use_container_width=True)
                    
            with colB:
                if 'Perfil_usuario__c' in df.columns and not df.empty:
                    admins = df.groupby('Perfil_usuario__c').size().reset_index(name='Operaciones')
                    fig_admins = px.bar(admins, x='Perfil_usuario__c', y='Operaciones', title='Operaciones por Administrativo', color='Operaciones')
                    st.plotly_chart(fig_admins, use_container_width=True)

        with tab2:
            st.subheader("Auditoría y Detalles de Operaciones")
            total_boletos = len(df)
            total_finalizados = len(df[df['Estado Actual'] == 'Disfruta Tu Nuevo Toyota'])
            total_bajas = len(df[df['Estado Actual'] == 'Operación Dada De Baja'])
            total_retrasados = retrasados_mask.sum()
            
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("📁 Total Boletos", total_boletos)
            c2.metric("✅ Finalizados", total_finalizados)
            c3.metric("❌ Bajas", total_bajas)
            c4.metric("⚠️ Operaciones con Retraso", total_retrasados)
            
            st.markdown("---")
            st.write("**Explorar boletos por estado específico:**")
            
            orden_filtro = ["Ingreso Pendiente", "Creacion En Siac", "Proboleto Aprobado", "Pedido Confirmado", "Firma Del Boleto", "En Proceso De Pago", "Facturacion", "Poliza De Seguro", "En Proceso De Patentamiento", "En Proceso De Entrega", "Disfruta Tu Nuevo Toyota", "Operación Dada De Baja"]
            estados_en_df = df['Estado Actual'].unique()
            estados_disponibles = [e for e in orden_filtro if e in estados_en_df] + [e for e in estados_en_df if e not in orden_filtro]
            
            estado_seleccionado = st.selectbox("Filtrar tabla por estado:", ["Todos los estados"] + estados_disponibles)
            
            df_tabla = df.copy() if estado_seleccionado == "Todos los estados" else df[df['Estado Actual'] == estado_seleccionado]
            
            def aplicar_estilos_dinamicos(row):
                styles = [''] * len(row)
                estado = row.get('Estado Actual', '')
                dias = row.get('Días en este estado', 0)
                
                if 'Estado Actual' in row.index:
                    idx_estado = row.index.get_loc('Estado Actual')
                    if estado == 'Disfruta Tu Nuevo Toyota':
                        styles[idx_estado] = 'background-color: rgba(46, 204, 113, 0.3); font-weight: bold;'
                    elif estado == 'Operación Dada De Baja':
                        styles[idx_estado] = 'background-color: rgba(149, 165, 166, 0.4); color: #7f8c8d; font-style: italic;'
                        
                if 'Días en este estado' in row.index and estado not in ['Disfruta Tu Nuevo Toyota', 'Operación Dada De Baja']:
                    if dias > limites_demora.get(estado, 5):
                        idx_dias = row.index.get_loc('Días en este estado')
                        styles[idx_dias] = 'background-color: rgba(255, 75, 75, 0.3); color: #c0392b; font-weight: bold;'
                        
                return styles

            columnas_mostrar = ['Identificador', 'Numero_de_Boleto__c', 'Sucursal_de_Venta__c', 'Denominacion_Comercial__c', 'Nombres_de_Cuenta__c', 'Estado Actual', 'Días en este estado', 'Fecha de Último Estado', 'Nombre_Asesor__c', 'Comentario']
            columnas_relevantes = [c for c in columnas_mostrar if c in df_tabla.columns]
            
            st.dataframe(
                df_tabla[columnas_relevantes].style.apply(aplicar_estilos_dinamicos, axis=1),
                use_container_width=True, hide_index=True
            )
            
            st.download_button("📥 Descargar tabla de Auditoría (Excel)", data=convertir_df_a_excel(df_tabla[columnas_relevantes]), file_name=f'Auditoria_Boletos.xlsx', mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

        with tab3:
            st.subheader("Línea de Tiempo por Cliente")
            if 'Identificador' in df.columns and not df.empty:
                seleccion = st.selectbox("Escribe o selecciona el nombre del cliente o boleto:", sorted(df['Identificador'].unique()), index=None, placeholder="Ej: PEREZ JUAN...")
                if seleccion:
                    datos_boleto = df[df['Identificador'] == seleccion].iloc[0]
                    col_info1, col_info2, col_info3 = st.columns(3)
                    col_info1.info(f"**Estado Actual:** {datos_boleto['Estado Actual']}")
                    
                    if datos_boleto['Estado Actual'] == "Disfruta Tu Nuevo Toyota":
                        col_info2.success("✅ **Operación finalizada.**")
                    elif datos_boleto['Estado Actual'] == "Operación Dada De Baja":
                        col_info2.error("❌ **Operación Cancelada.**")
                    else:
                        dias_actuales = datos_boleto['Días en este estado']
                        limite = limites_demora.get(datos_boleto['Estado Actual'], 5)
                        if dias_actuales > limite: col_info2.error(f"⚠️ **RETRASADO:** {dias_actuales} días estancado.")
                        else: col_info2.info(f"⏳ **A tiempo:** Lleva {dias_actuales} días.")
                    
                    if 'Fecha_de_entrega_estimada__c' in df.columns and pd.notna(datos_boleto['Fecha_de_entrega_estimada__c']):
                        col_info3.metric("📅 Entrega Estimada", datos_boleto['Fecha_de_entrega_estimada__c'].strftime('%d-%b-%Y'))
                    else:
                        col_info3.metric("📅 Entrega Estimada", "No definida")
                    
                    hitos, fechas = [], []
                    for col in cols_fechas:
                        if col != 'Fecha_de_entrega_estimada__c' and pd.notna(datos_boleto[col]):
                            hitos.append(str(col).replace('Fecha_de_', '').replace('Fecha_', '').replace('__c', '').replace('_', ' ').title())
                            fechas.append(datos_boleto[col])
                    
                    if hitos:
                        df_timeline = pd.DataFrame({'Etapa': hitos, 'Fecha': fechas})
                        orden_ideal = ["Creacion En Siac", "Proboleto Aprobado", "Pedido Confirmado", "Firma Del Boleto", "En Proceso De Pago", "Facturacion", "Poliza De Seguro", "En Proceso De Patentamiento", "En Proceso De Entrega", "Disfruta Tu Nuevo Toyota", "Baja"]
                        df_timeline['Etapa'] = pd.Categorical(df_timeline['Etapa'], categories=orden_ideal, ordered=True)
                        df_timeline = df_timeline.sort_values(by='Etapa')
                        
                        fig_timeline = px.line(df_timeline, x='Fecha', y='Etapa', markers=True, title=f"Historial: {seleccion}", text='Fecha')
                        fig_timeline.update_traces(line_color='#2c3e50', marker=dict(size=12, color='#e74c3c'), textposition="top center", texttemplate='%{text|%d-%b-%Y}')
                        fig_timeline.update_yaxes(categoryorder='array', categoryarray=orden_ideal, autorange="reversed")
                        st.plotly_chart(fig_timeline, use_container_width=True)
                    else: st.info("Boleto sin fechas registradas.")
            else: st.info("No se encontraron las columnas de Cliente o la tabla está vacía.")

        with tab4:
            st.subheader("Visión General del Flujo")
            col1, col2, col3 = st.columns(3)
            col1.metric("Operaciones (Boletos)", len(df))
            vehiculos_pendientes = df[(df['Estado Actual'] != 'Disfruta Tu Nuevo Toyota') & (df['Estado Actual'] != 'Operación Dada De Baja')]
            demora_media = vehiculos_pendientes['Días en este estado'].mean() if not vehiculos_pendientes.empty else 0
            col2.metric("Demora Promedio Actual", f"{demora_media:.1f} días" if pd.notna(demora_media) else "0 días")
            
            st.markdown("---")
            st.subheader("Cuellos de Botella: Boletos en cada etapa")
            lista_etapas = ["Ingreso Pendiente", "Creacion En Siac", "Proboleto Aprobado", "Pedido Confirmado", "Firma Del Boleto", "En Proceso De Pago", "Facturacion", "Poliza De Seguro", "En Proceso De Patentamiento", "En Proceso De Entrega", "Disfruta Tu Nuevo Toyota", "Operación Dada De Baja"]
            conteos_estado = df['Estado Actual'].value_counts().to_dict()
            df_cuellos = pd.DataFrame([{'Etapa': e, 'Boletos Detenidos': conteos_estado.get(e, 0)} for e in reversed(lista_etapas)])
            
            fig_cuellos = px.bar(df_cuellos, x='Boletos Detenidos', y='Etapa', orientation='h', title='Volumen estancado por proceso', color='Boletos Detenidos', color_continuous_scale=['#f0f2f6', '#ff4b4b'])
            fig_cuellos.update_layout(showlegend=False)
            st.plotly_chart(fig_cuellos, use_container_width=True)
            
            st.markdown("---")
            if not vehiculos_pendientes.empty:
                historico = vehiculos_pendientes.groupby('Mes')['Días en este estado'].mean().reset_index()
                historico = historico[historico['Mes'] != 'NaT'].sort_values(by='Mes')
                if not historico.empty:
                    fig_linea = px.line(historico, x='Mes', y='Días en este estado', title="Evolución Histórica de Demoras", markers=True)
                    fig_linea.update_traces(line_color='red') 
                    st.plotly_chart(fig_linea, use_container_width=True)
                
        with tab5:
            st.subheader("⚠️ Registro de Boletos Demorados")
            df_demorados = df[retrasados_mask].copy()
            
            if not df_demorados.empty:
                df_demorados = df_demorados.sort_values(by='Días en este estado', ascending=False)
                columnas_demora = ['Identificador', 'Numero_de_Boleto__c', 'Estado Actual', 'Días en este estado', 'Nombre_Asesor__c', 'Comentario']
                df_demorados_mostrar = df_demorados[[c for c in columnas_demora if c in df_demorados.columns]]
                
                df_editado_demorados = st.data_editor(
                    df_demorados_mostrar,
                    use_container_width=True, hide_index=True,
                    column_config={"Identificador": None, "Comentario": st.column_config.TextColumn("💬 Comentario (Doble clic)", help="Escribe el motivo de la demora.")},
                    disabled=[c for c in columnas_demora if c != 'Comentario'], key="editor_tabla_demoras"
                )
                
                if df_editado_demorados is not None:
                    for idx in df_editado_demorados.index:
                        df.loc[df['Identificador'] == df_editado_demorados.at[idx, 'Identificador'], 'Comentario'] = df_editado_demorados.at[idx, 'Comentario']
                    df[['Identificador', 'Comentario']].drop_duplicates(subset=['Identificador']).to_csv(ARCHIVO_COMENTARIOS, index=False)
                
                st.download_button("📥 Descargar tabla de Demoras (Excel)", data=convertir_df_a_excel(df_editado_demorados), file_name='Alertas_de_Demora.xlsx', mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            else:
                st.success("¡Excelente! No hay boletos demorados en este momento.")

    except Exception as e:
        st.error(f"Error procesando la base de datos. (Detalle: {e})")
