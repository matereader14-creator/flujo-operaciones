import os
import io
import streamlit as st
import pandas as pd
import plotly.express as px
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

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

@st.cache_data(ttl=600)
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
    st.sidebar.warning("Link de Google Sheets no configurado.")

# ==========================================
# DIRECTORIO DE CORREOS
# ==========================================
diccionario_correos = {
    "Ernesto José Luis Salomón": "ernesto.salomon@autolux.com.ar",
    "Nicolas Scachi": "nicolas.scacchi@autolux.com.ar",
    "Gabriel Lopez Quiroga": "gabriel.quiroga@autolux.com.ar",
    "Matias Bravo": "matias.bravo@autolux.com.ar",
    "Jorge Agustín Gonzalez": "jorge.gonzalez@autolux.com.ar",
    "Franco Oscar Lesser": "franco.lesser@autolux.com.ar",
    "Vilma Viviana Carabajal": "vilma.carabajal@autolux.com.ar",
    "Arnaldo Denis Peralta": "arnaldo.peralta@autolux.com.ar",
    "Nicolas Elbio Cordoba": "nicolas.cordoba@autolux.com.ar",
    "María Cecilia Fernandez": "cecilia.fernandez@autolux.com.ar",
    "Nadia Macar Fernandez Colletti": "nadia.fernandez@autolux.com.ar",
    "Facundo Gabriel Ponce Cavion": "facundo.ponce@autolux.com.ar",
    "Pablo Carrizo": "pablo.carrizo@cenoa.com.ar",
    "Romina Romagnoli": "romina.romagnoli@autolux.com.ar",
    "Guillermo Nallim": "guillermo.nallim@autolux.com.ar",
    "Mauro Aramayo": "mauro.aramayo@autolux.com.ar",
    "Luis Humano": "luis.humano@autolux.com.ar",
    "Gonzalo Daniel Ponce Cavion": "gonzalo.ponce@autolux.com.ar",
    "Gustavo Enrique Arias Mazza": "gustavo.arias@autolux.com.ar",
    "Jose Amoros": "jose.amoros@autolux.com.ar",
    "Gustavo Cabezas": "gustavo.cabezas@autolux.com.ar",
    "Jose Siñanez": "jose.siñanez@autolux.com.ar",
    "Soledad Zamora": "soledad.zamora@autolux.com.ar",
    "Cristian Noblega": "cristian.noblega@autolux.com.ar",
    "Faustino Ezequiel Cardozo": "ezequiel.cardozo@autolux.com.ar",
    "Daniel Flores": "daniel.flores@autolux.com.ar",
    "Gonzalo Martinez": "gonzalo.martinez@autolux.com.ar",
    "Leonel Mamani": "leonel.mamani@autolux.com.ar",
    "Cecilia Diaz": "cecilia.diaz@autolux.com.ar",
    "Martín Andrés Diaz": "martin.diaz@autolux.com.ar"
}

# --- CLASIFICACIÓN DE ADMINISTRADORES POR SUCURSAL ---
correos_administradores = {
    "GLOBAL": {
        # Espacio para direcciones que deban recibir copias de TODAS las sucursales (gerencias, etc.)
    },
    "SALTA": {
        "Franco Gallardo": "ventas.especiales@autolux.com.ar",
        "Mariano (Créditos)": "creditos.salta@autolux.com.ar",
        "Alfoncina Gianibelli": "alfonsina.gianibelli@autolux.com.ar",
        "Eugenia Alvarez": "maria.alvarez@autolux.com.ar"
    },
    "JUJUY": {
        "Luis Ledesma": "luis.ledesma@autolux.com.ar",
        "Mariano Walker": "mariano.walker@autolux.com.ar",
        "Karen Diaz": "karen.diaz@autolux.com.ar"
    },
    "TARTAGAL": {
        "Ramiro Durand": "ramiro.durand@autolux.com.ar"
    }
}

# ==========================================
# FUNCIÓN PARA ENVIAR CORREOS CRUZADOS
# ==========================================
def enviar_correo_personalizado(df_alertas, destinatario, cc, asunto, saludo, intro_texto):
    try:
        remitente = st.secrets["EMAIL_REMITENTE"]
        password = st.secrets["EMAIL_PASSWORD"]
    except Exception:
        return False, "Falta configurar las contraseñas en Streamlit Secrets."

    msg = MIMEMultipart()
    msg['From'] = remitente
    msg['To'] = destinatario
    if cc:
        msg['Cc'] = cc
    msg['Subject'] = asunto
    
    html_table = df_alertas.to_html(index=False, border=0, justify='center')
    html_table = html_table.replace('<table', '<table style="border-collapse: collapse; width: 100%; font-family: Arial, sans-serif;"')
    html_table = html_table.replace('<th>', '<th style="background-color: #d32f2f; color: white; padding: 10px; border: 1px solid #ddd;">')
    html_table = html_table.replace('<td>', '<td style="padding: 10px; border: 1px solid #ddd; text-align: center;">')
    
    html_body = f"""
    <html>
    <body>
        <h2 style="color: #d32f2f; font-family: Arial, sans-serif;">Reporte de Boletos con Retraso Operativo</h2>
        <p style="font-family: Arial, sans-serif; font-size: 14px;">{saludo}</p>
        <p style="font-family: Arial, sans-serif; font-size: 14px;">{intro_texto}</p>
        {html_table}
        <br>
        <p style="font-family: Arial, sans-serif; font-size: 12px; color: #777;">Este es un mensaje automático generado por la Plataforma de Seguimiento de Operaciones de Calidad LUX.</p>
    </body>
    </html>
    """
    msg.attach(MIMEText(html_body, 'html'))
    
    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(remitente, password)
        server.send_message(msg)
        server.quit()
        return True, ""
    except Exception as e:
        return False, str(e)

# ==========================================
# 3. PROCESAMIENTO Y GRÁFICOS
# ==========================================
if df_cargado is not None:
    try:
        df = df_cargado.copy()
        
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
            
        ARCHIVO_COMENTARIOS = 'comentarios_operaciones.csv'
        if os.path.exists(ARCHIVO_COMENTARIOS):
            df_com = pd.read_csv(ARCHIVO_COMENTARIOS)
            df = df.merge(df_com, on='Identificador', how='left')
        else:
            df['Comentario'] = ""
            
        df['Comentario'] = df['Comentario'].fillna("")
        
        # PROCESAMIENTO DE FECHAS
        cols_fechas = [c for c in df.columns if 'fecha' in str(c).lower()]
        for col in cols_fechas:
            df[col] = pd.to_datetime(df[col], errors='coerce')
            
        cols_fechas_desarrollo = [c for c in cols_fechas if 'entrega_estimada' not in str(c).lower()]
        cols_fechas_calculo_inicio = [c for c in cols_fechas_desarrollo if 'baja' not in str(c).lower()]
        
        df = df.dropna(subset=cols_fechas_desarrollo, how='all')
        
        # Filtro estricto de año 2026
        mask_2026 = pd.Series(False, index=df.index)
        for col in cols_fechas_desarrollo:
            mask_2026 = mask_2026 | (df[col].dt.year == 2026)
        
        df = df[mask_2026].copy()
        
        # Limites de demora
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
                    for col in cols_fechas_desarrollo:
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
            
            limite = limites_demora.get(estado_actual, "-")
                
            return pd.Series([estado_actual, demora, limite, ultima_fecha])
            
        if df.empty:
            df['Estado Actual'] = pd.Series(dtype='object')
            df['Días en este estado'] = pd.Series(dtype='float64')
            df['Límite Máximo'] = pd.Series(dtype='object')
            df['Fecha de Último Estado'] = pd.Series(dtype='datetime64[ns]')
            df['Mes'] = pd.Series(dtype='object')
        else:
            df[['Estado Actual', 'Días en este estado', 'Límite Máximo', 'Fecha de Último Estado']] = df.apply(calcular_estado_y_demora, axis=1)
            df['Mes'] = df['Fecha de Último Estado'].dt.to_period('M').astype(str)
        
        if df.empty:
            retrasados_mask = pd.Series(False, index=df.index)
            amarillos_mask = pd.Series(False, index=df.index)
        else:
            retrasados_mask = df.apply(
                lambda x: x['Días en este estado'] > limites_demora.get(x['Estado Actual'], 5) 
                if x['Estado Actual'] not in ['Disfruta Tu Nuevo Toyota', 'Operación Dada De Baja'] else False, 
                axis=1
            )
            amarillos_mask = df.apply(
                lambda x: (x['Días en este estado'] >= limites_demora.get(x['Estado Actual'], 5) * 0.8) and 
                          (x['Días en este estado'] <= limites_demora.get(x['Estado Actual'], 5))
                if x['Estado Actual'] not in ['Disfruta Tu Nuevo Toyota', 'Operación Dada De Baja'] else False,
                axis=1
            )
        
        st.sidebar.markdown("---")
        st.sidebar.subheader("Filtros Globales")
        
        meses_limpios = [str(m) for m in df['Mes'].unique() if str(m) not in ('NaT', 'nan', 'None')]
        meses_disponibles = sorted(meses_limpios)
        opciones_mes = ["Todos los meses"] + meses_disponibles
        
        # --- MES PREDETERMINADO AUTOMÁTICO (NATIVO, SIN PYTZ) ---
        mes_actual = (pd.Timestamp.utcnow() - pd.Timedelta(hours=3)).strftime('%Y-%m') 
        
        idx_mes_defecto = 0
        if mes_actual in opciones_mes:
            idx_mes_defecto = opciones_mes.index(mes_actual)
            
        mes_seleccionado = st.sidebar.selectbox("Seleccionar Mes:", opciones_mes, index=idx_mes_defecto)
        
        if mes_seleccionado != "Todos los meses":
            df = df[df['Mes'] == mes_seleccionado].copy()
            retrasados_mask = df.apply(
                lambda x: x['Días en este estado'] > limites_demora.get(x['Estado Actual'], 5) 
                if x['Estado Actual'] not in ['Disfruta Tu Nuevo Toyota', 'Operación Dada De Baja'] else False, 
                axis=1
            )
            amarillos_mask = df.apply(
                lambda x: (x['Días en este estado'] >= limites_demora.get(x['Estado Actual'], 5) * 0.8) and 
                          (x['Días en este estado'] <= limites_demora.get(x['Estado Actual'], 5))
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
                    amarillos_mask = df.apply(
                        lambda x: (x['Días en este estado'] >= limites_demora.get(x['Estado Actual'], 5) * 0.8) and 
                                  (x['Días en este estado'] <= limites_demora.get(x['Estado Actual'], 5))
                        if x['Estado Actual'] not in ['Disfruta Tu Nuevo Toyota', 'Operación Dada De Baja'] else False,
                        axis=1
                    )
        
        def convertir_df_a_excel(df_export):
            salida = io.BytesIO()
            with pd.ExcelWriter(salida, engine='openpyxl') as writer:
                df_export.to_excel(writer, index=False, sheet_name='Datos_Boletos')
            return salida.getvalue()

        # ==========================================
        # CREACIÓN DE PESTAÑAS
        # ==========================================
        tab_auditoria, tab_rendimiento, tab_rastreo, tab_flujo, tab_alertas = st.tabs([
            "🔍 Auditoría",
            "👥 Rendimiento", 
            "🔎 Rastreo",
            "📊 Flujo",
            "⚠️ Alertas de Demora"
        ])
        
        with tab_auditoria:
            st.subheader("Auditoría y Detalles de Operaciones")
            total_boletos = len(df)
            total_finalizados = len(df[df['Estado Actual'] == 'Disfruta Tu Nuevo Toyota'])
            total_bajas = len(df[df['Estado Actual'] == 'Operación Dada De Baja'])
            total_amarillos = amarillos_mask.sum()
            total_retrasados = retrasados_mask.sum()
            
            c1, c2, c3, c4, c5 = st.columns(5)
            c1.metric("📁 Total Boletos", total_boletos)
            c2.metric("✅ Finalizados", total_finalizados)
            c3.metric("❌ Bajas", total_bajas)
            c4.metric("🟡 Alerta Preventiva", total_amarillos)
            c5.metric("⚠️ Retrasados", total_retrasados)
            
            st.markdown("---")
            st.write("**Explorar boletos por estado específico:**")
            
            orden_filtro = ["Ingreso Pendiente", "Creacion En Siac", "Proboleto Aprobado", "Pedido Confirmado", "Firma Del Boleto", "En Proceso De Pago", "Facturacion", "Poliza De Seguro", "En Proceso De Patentamiento", "En Proceso De Entrega", "Disfruta Tu Nuevo Toyota", "Operación Dada De Baja"]
            estados_en_df = df['Estado Actual'].unique()
            estados_disponibles = [e for e in orden_filtro if e in estados_en_df] + [e for e in estados_en_df if e not in orden_filtro]
            
            estado_seleccionado = st.selectbox("Filtrar tabla por estado:", ["Todos los estados"] + estados_disponibles)
            
            df_tabla = df.copy() if estado_seleccionado == "Todos los estados" else df[df['Estado Actual'] == estado_seleccionado]
            
            # Lógica Bajas al Final
            df_tabla['es_baja'] = df_tabla['Estado Actual'] == 'Operación Dada De Baja'
            df_tabla = df_tabla.sort_values(by=['es_baja', 'Días en este estado'], ascending=[True, False]).drop(columns=['es_baja'])
            
            def aplicar_estilos_dinamicos(row):
                styles = [''] * len(row)
                estado = row.get('Estado Actual', '')
                dias = row.get('Días en este estado', 0)
                limite = limites_demora.get(estado, 5)
                
                if 'Estado Actual' in row.index:
                    idx_estado = row.index.get_loc('Estado Actual')
                    if estado == 'Disfruta Tu Nuevo Toyota':
                        styles[idx_estado] = 'background-color: rgba(46, 204, 113, 0.3); font-weight: bold;'
                    elif estado == 'Operación Dada De Baja':
                        styles[idx_estado] = 'background-color: rgba(149, 165, 166, 0.4); color: #7f8c8d; font-style: italic;'
                        
                if 'Días en este estado' in row.index and estado not in ['Disfruta Tu Nuevo Toyota', 'Operación Dada De Baja']:
                    idx_dias = row.index.get_loc('Días en este estado')
                    if dias > limite:
                        styles[idx_dias] = 'background-color: rgba(255, 75, 75, 0.3); color: #c0392b; font-weight: bold;'
                    elif dias >= limite * 0.8:
                        styles[idx_dias] = 'background-color: rgba(241, 196, 15, 0.3); color: #d68910; font-weight: bold;'
                        
                return styles

            columnas_mostrar = ['Identificador', 'Numero_de_Boleto__c', 'Sucursal_de_Venta__c', 'Denominacion_Comercial__c', 'Nombres_de_Cuenta__c', 'Estado Actual', 'Días en este estado', 'Límite Máximo', 'Fecha de Último Estado', 'Nombre_Asesor__c', 'Comentario']
            columnas_relevantes = [c for c in columnas_mostrar if c in df_tabla.columns]
            
            st.dataframe(
                df_tabla[columnas_relevantes].style.apply(aplicar_estilos_dinamicos, axis=1),
                use_container_width=True, hide_index=True
            )
            
            st.download_button("📥 Descargar tabla de Auditoría (Excel)", data=convertir_df_a_excel(df_tabla[columnas_relevantes]), file_name=f'Auditoria_Boletos.xlsx', mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

        with tab_rendimiento:
            st.subheader("Carga de Trabajo Operativo - Asesores")
            if 'Nombre_Asesor__c' in df.columns and not df.empty:
                vendedores = df.groupby('Nombre_Asesor__c').size().reset_index(name='Operaciones')
                vendedores = vendedores.sort_values(by='Operaciones', ascending=False)
                
                fig_vendedores = px.bar(
                    vendedores, 
                    x='Nombre_Asesor__c', 
                    y='Operaciones', 
                    title='Operaciones por Asesor', 
                    color='Nombre_Asesor__c', 
                    color_discrete_sequence=px.colors.qualitative.Dark24
                )
                fig_vendedores.update_layout(showlegend=False)
                st.plotly_chart(fig_vendedores, use_container_width=True)

        with tab_rastreo:
            st.subheader("Línea de Tiempo por Cliente")
            if 'Identificador' in df.columns and not df.empty:
                seleccion = st.selectbox("Escribe o selecciona el nombre del cliente o boleto:", sorted(df['Identificador'].unique()), index=None, placeholder="Ej: PEREZ JUAN...")
                if seleccion:
                    datos_boleto = df[df['Identificador'] == seleccion].iloc[0]
                    col_info1, col_info2, col_info3, col_info4 = st.columns(4)
                    col_info1.info(f"**Estado Actual:** {datos_boleto['Estado Actual']}")
                    
                    if datos_boleto['Estado Actual'] == "Disfruta Tu Nuevo Toyota":
                        col_info2.success("✅ **Operación finalizada.**")
                    elif datos_boleto['Estado Actual'] == "Operación Dada De Baja":
                        col_info2.error("❌ **Operación Cancelada.**")
                    else:
                        dias_actuales = datos_boleto['Días en este estado']
                        limite = limites_demora.get(datos_boleto['Estado Actual'], 5)
                        if dias_actuales > limite: col_info2.error(f"⚠️ **RETRASADO:** {dias_actuales} días estancado.")
                        elif dias_actuales >= limite * 0.8: col_info2.warning(f"🟡 **PRECAUCIÓN:** {dias_actuales}/{limite} días consumidos.")
                        else: col_info2.info(f"⏳ **A tiempo:** Lleva {dias_actuales} días.")
                    
                    if 'Fecha_de_entrega_estimada__c' in df.columns and pd.notna(datos_boleto['Fecha_de_entrega_estimada__c']):
                        col_info3.metric("📅 Entrega Estimada", datos_boleto['Fecha_de_entrega_estimada__c'].strftime('%d-%b-%Y'))
                    else:
                        col_info3.metric("📅 Entrega Estimada", "No definida")
                        
                    fechas_inicio_validas = pd.to_datetime(datos_boleto[cols_fechas_calculo_inicio].dropna(), errors='coerce')
                    fecha_inicio_boleto = fechas_inicio_validas.min() if not fechas_inicio_validas.empty else pd.NaT
                    
                    if datos_boleto['Estado Actual'] == 'Disfruta Tu Nuevo Toyota':
                        fecha_disp = datos_boleto.get('Fecha_de_disfruta_tu_nuevo_Toyota__c')
                        if pd.notna(fecha_disp):
                            fecha_fin_boleto = pd.to_datetime(fecha_disp)
                        else:
                            fechas_fin_validas = pd.to_datetime(datos_boleto[cols_fechas_desarrollo].dropna(), errors='coerce')
                            fecha_fin_boleto = fechas_fin_validas.max() if not fechas_fin_validas.empty else pd.NaT
                        texto_duracion = "⏱️ Duración Total"
                        
                    elif datos_boleto['Estado Actual'] == 'Operación Dada De Baja':
                        fecha_baja = datos_boleto.get('Fecha_de_baja__c')
                        if pd.notna(fecha_baja):
                            fecha_fin_boleto = pd.to_datetime(fecha_baja)
                        else:
                            fechas_fin_validas = pd.to_datetime(datos_boleto[cols_fechas_desarrollo].dropna(), errors='coerce')
                            fecha_fin_boleto = fechas_fin_validas.max() if not fechas_fin_validas.empty else pd.NaT
                        texto_duracion = "⏱️ Duración hasta Baja"
                        
                    else:
                        fecha_fin_boleto = pd.Timestamp.now()
                        texto_duracion = "⏱️ Días Transcurridos"
                        
                    if pd.notna(fecha_inicio_boleto) and pd.notna(fecha_fin_boleto):
                        dias_totales_boleto = (fecha_fin_boleto - fecha_inicio_boleto).days
                        col_info4.metric(texto_duracion, f"{max(0, dias_totales_boleto)} días")
                    else:
                        col_info4.metric(texto_duracion, "No calculable")
                    
                    hitos, fechas = [], []
                    for col in cols_fechas:
                        if col != 'Fecha_de_entrega_estimada__c' and pd.notna(datos_boleto[col]):
                            fecha_valida = pd.to_datetime(datos_boleto[col], errors='coerce')
                            if pd.notna(fecha_valida):
                                hitos.append(str(col).replace('Fecha_de_', '').replace('Fecha_', '').replace('__c', '').replace('_', ' ').title())
                                fechas.append(fecha_valida)
                    
                    if hitos:
                        df_timeline = pd.DataFrame({'Etapa': hitos, 'Fecha': fechas})
                        orden_ideal = ["Creacion En Siac", "Proboleto Aprobado", "Pedido Confirmado", "Firma Del Boleto", "En Proceso De Pago", "Facturacion", "Poliza De Seguro", "En Proceso De Patentamiento", "En Proceso De Entrega", "Disfruta Tu Nuevo Toyota", "Baja"]
                        df_timeline['Etapa'] = pd.Categorical(df_timeline['Etapa'], categories=orden_ideal, ordered=True)
                        df_timeline = df_timeline.sort_values(by='Etapa')
                        
                        fig_timeline = px.line(df_timeline, x='Fecha', y='Etapa', markers=True, title=f"Historial: {seleccion}", text='Fecha')
                        fig_timeline.update_traces(line_color='#2c3e50', marker=dict(size=12, color='#e74c3c'), textposition="top center", texttemplate='%{text|%d-%b-%Y}')
                        fig_timeline.update_yaxes(categoryorder='array', categoryarray=orden_ideal, autorange="reversed")
                        st.plotly_chart(fig_timeline, use_container_width=True)

        with tab_flujo:
            st.subheader("Visión General del Flujo")
            col1, col2, col3 = st.columns(3)
            col1.metric("Operaciones (Boletos)", len(df))
            vehiculos_pendientes = df[(df['Estado Actual'] != 'Disfruta Tu Nuevo Toyota') & (df['Estado Actual'] != 'Operación Dada De Baja')]
            demora_media = vehiculos_pendientes['Días en este estado'].mean() if not vehiculos_pendientes.empty else 0
            col2.metric("Demora Promedio Actual", f"{demora_media:.1f} días" if pd.notna(demora_media) else "0 días")
            
            df_finalizados_historico = df[df['Estado Actual'] == 'Disfruta Tu Nuevo Toyota'].copy()
            if not df_finalizados_historico.empty:
                inicios_historico = df_finalizados_historico[cols_fechas_calculo_inicio].min(axis=1)
                
                if 'Fecha_de_disfruta_tu_nuevo_Toyota__c' in df_finalizados_historico.columns:
                    fines_historico = df_finalizados_historico['Fecha_de_disfruta_tu_nuevo_Toyota__c']
                else:
                    fines_historico = df_finalizados_historico[cols_fechas_desarrollo].max(axis=1)
                    
                tiempos_finalizacion = (pd.to_datetime(fines_historico, errors='coerce') - pd.to_datetime(inicios_historico, errors='coerce')).dt.days
                promedio_finalizacion = tiempos_finalizacion.mean()
                col3.metric("Tiempo Promedio de Finalización", f"{promedio_finalizacion:.1f} días" if pd.notna(promedio_finalizacion) else "N/A")
            else:
                col3.metric("Tiempo Promedio de Finalización", "0 días")
            
            st.markdown("---")
            st.subheader("Cuellos de Botella: Boletos Activos en cada etapa")
            
            lista_etapas = ["Ingreso Pendiente", "Creacion En Siac", "Proboleto Aprobado", "Pedido Confirmado", "Firma Del Boleto", "En Proceso De Pago", "Facturacion", "Poliza De Seguro", "En Proceso De Patentamiento", "En Proceso De Entrega", "Disfruta Tu Nuevo Toyota"]
            conteos_estado = df['Estado Actual'].value_counts().to_dict()
            df_cuellos = pd.DataFrame([{'Etapa': e, 'Boletos Detenidos': conteos_estado.get(e, 0)} for e in reversed(lista_etapas)])
            
            fig_cuellos = px.bar(df_cuellos, x='Boletos Detenidos', y='Etapa', orientation='h', title='Volumen estancado por proceso (Sin Bajas)', color='Boletos Detenidos', color_continuous_scale=['#f0f2f6', '#ff4b4b'])
            fig_cuellos.update_layout(showlegend=False)
            st.plotly_chart(fig_cuellos, use_container_width=True)
            
            st.markdown("---")
            st.subheader("❌ Análisis de Operaciones Canceladas (Bajas)")
            total_bajas_flujo = len(df[df['Estado Actual'] == 'Operación Dada De Baja'])
            porcentaje_bajas = (total_bajas_flujo / len(df) * 100) if len(df) > 0 else 0
            
            st.metric(
                label="Total de Bajas en el mes seleccionado", 
                value=f"{total_bajas_flujo} boletos", 
                delta=f"{porcentaje_bajas:.1f}% del volumen total", 
                delta_color="inverse"
            )
            st.info("Nota: Las operaciones dadas de baja han sido separadas del flujo principal para no alterar la lectura del embudo comercial activo.")
                
        with tab_alertas:
            st.subheader("⚠️ Registro de Boletos Demorados")
            df_demorados = df[retrasados_mask].copy()
            
            if not df_demorados.empty:
                df_demorados = df_demorados.sort_values(by='Días en este estado', ascending=False)
                columnas_demora = ['Identificador', 'Numero_de_Boleto__c', 'Estado Actual', 'Días en este estado', 'Límite Máximo', 'Nombre_Asesor__c', 'Comentario']
                df_demorados_mostrar = df_demorados[[c for c in columnas_demora if c in df_demorados.columns]]
                
                busqueda = st.text_input("🔍 Buscar en demoras (Por Asesor, Boleto, Cliente, Estado...):", placeholder="Ej: Perez, 6684715, Confirmado...")
                if busqueda:
                    mask_busqueda = df_demorados_mostrar.astype(str).apply(lambda x: x.str.contains(busqueda, case=False, na=False)).any(axis=1)
                    df_demorados_mostrar = df_demorados_mostrar[mask_busqueda]
                
                df_editado_demorados = st.data_editor(
                    df_demorados_mostrar,
                    use_container_width=True, hide_index=True,
                    column_config={"Identificador": None, "Comentario": st.column_config.TextColumn("💬 Comentario", help="Escribe el motivo.")},
                    disabled=[c for c in columnas_demora if c != 'Comentario'], key="editor_demoras_v21"
                )
                
                if df_editado_demorados is not None:
                    for idx in df_editado_demorados.index:
                        df.loc[df['Identificador'] == df_editado_demorados.at[idx, 'Identificador'], 'Comentario'] = df_editado_demorados.at[idx, 'Comentario']
                    df[['Identificador', 'Comentario']].drop_duplicates(subset=['Identificador']).to_csv(ARCHIVO_COMENTARIOS, index=False)
                
                if 'Nombre_Asesor__c' in df_editado_demorados.columns and not df_editado_demorados.empty:
                    st.markdown("---")
                    
                    conteo = df_editado_demorados['Nombre_Asesor__c'].value_counts().reset_index()
                    conteo.columns = ['Asesor', 'Cantidad']
                    
                    total_demoras = conteo['Cantidad'].sum()
                    fila_total = pd.DataFrame([{'Asesor': 'TOTAL', 'Cantidad': total_demoras}])
                    
                    conteo_normal = conteo.sort_values(by='Cantidad', ascending=True)
                    conteo_final = pd.concat([conteo_normal, fila_total], ignore_index=True)
                    
                    mapa_colores = {asesor: px.colors.qualitative.Dark24[i % 24] for i, asesor in enumerate(conteo_normal['Asesor'])}
                    mapa_colores['TOTAL'] = '#d32f2f'
                    
                    fig_resumen = px.bar(
                        conteo_final, 
                        x='Cantidad', 
                        y='Asesor', 
                        orientation='h',
                        text='Cantidad',
                        title='📊 Resumen de Boletos Demorados por Asesor',
                        color='Asesor',
                        color_discrete_map=mapa_colores
                    )
                    
                    fig_resumen.update_traces(textposition='auto')
                    fig_resumen.update_layout(showlegend=False, xaxis_title="Cantidad de Boletos Detenidos", yaxis_title="")
                    
                    st.plotly_chart(fig_resumen, use_container_width=True)
                
                st.markdown("---")
                col_b1, col_b2 = st.columns([1, 1])
                with col_b1:
                    st.download_button("📥 Descargar (Excel)", data=convertir_df_a_excel(df_editado_demorados), file_name='Demoras.xlsx', mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
                
                with col_b2:
                    if st.button("📧 Enviar Alertas por Etapa del Proceso"):
                        with st.spinner("Clasificando etapas y enviando correos cruzados por sucursal..."):
                            
                            df_mail_completo = df_editado_demorados.merge(df[['Identificador', 'Sucursal_de_Venta__c']], on='Identificador', how='left')
                            
                            if 'Nombre_Asesor__c' in df_mail_completo.columns:
                                estados_pre_facturacion = ["Ingreso Pendiente", "Creacion En Siac", "Proboleto Aprobado", "Pedido Confirmado", "Firma Del Boleto", "En Proceso De Pago"]
                                estados_post_facturacion = ["Facturacion", "Poliza De Seguro", "En Proceso De Patentamiento", "En Proceso De Entrega"]
                                
                                df_pre = df_mail_completo[df_mail_completo['Estado Actual'].isin(estados_pre_facturacion)]
                                df_post = df_mail_completo[df_mail_completo['Estado Actual'].isin(estados_post_facturacion)]
                                
                                correos_enviados = 0
                                asesores_sin_correo = []
                                
                                def obtener_correos_admins(sucursal_str):
                                    suc_texto = str(sucursal_str).lower()
                                    lista_admins = []
                                    lista_admins.extend(correos_administradores.get("GLOBAL", {}).values())
                                    
                                    if "tartagal" in suc_texto:
                                        lista_admins.extend(correos_administradores.get("TARTAGAL", {}).values())
                                    elif "jujuy" in suc_texto:
                                        lista_admins.extend(correos_administradores.get("JUJUY", {}).values())
                                    else:
                                        lista_admins.extend(correos_administradores.get("SALTA", {}).values())
                                        
                                    return ", ".join(list(set(lista_admins)))
                                
                                # 1. BLOQUE PRE-FACTURACIÓN
                                if not df_pre.empty:
                                    for (asesor, sucursal), df_grupo in df_pre.groupby(['Nombre_Asesor__c', 'Sucursal_de_Venta__c']):
                                        df_asesor = df_grupo.drop(columns=['Nombre_Asesor__c', 'Sucursal_de_Venta__c'])
                                        correo_asesor = diccionario_correos.get(str(asesor).strip(), st.secrets["EMAIL_DESTINO"])
                                        
                                        if correo_asesor == st.secrets["EMAIL_DESTINO"]:
                                            asesores_sin_correo.append(asesor)
                                            
                                        correos_admin_cc = obtener_correos_admins(sucursal)
                                        
                                        asunto_pre = f"⚠️ Demoras Comerciales (PRE-Facturación) - Asesor: {asesor} - Sucursal: {sucursal}"
                                        saludo_pre = f"Hola <b>{asesor}</b>,"
                                        intro_pre = "El sistema detectó las siguientes demoras en las etapas previas a la Facturación. Al estar en la fase inicial del proceso, solicitamos tu gestión comercial para destrabar estas operaciones. <b>Por favor, ingresa a la aplicación de Flujo de Operaciones y, en la pestaña de 'Alertas de Demora', completa la casilla de 'Comentario' para informar el estado actual de cada caso:</b>"
                                        
                                        exito, _ = enviar_correo_personalizado(df_asesor, destinatario=correo_asesor, cc=correos_admin_cc, asunto=asunto_pre, saludo=saludo_pre, intro_texto=intro_pre)
                                        if exito: correos_enviados += 1
                                
                                # 2. BLOQUE POST-FACTURACIÓN
                                if not df_post.empty:
                                    for (asesor, sucursal), df_grupo in df_post.groupby(['Nombre_Asesor__c', 'Sucursal_de_Venta__c']):
                                        df_asesor = df_grupo.drop(columns=['Nombre_Asesor__c', 'Sucursal_de_Venta__c'])
                                        correo_asesor = diccionario_correos.get(str(asesor).strip(), st.secrets["EMAIL_DESTINO"])
                                        
                                        if correo_asesor == st.secrets["EMAIL_DESTINO"] and asesor not in asesores_sin_correo:
                                            asesores_sin_correo.append(asesor)
                                            
                                        correos_admin_destinatario = obtener_correos_admins(sucursal)
                                        
                                        asunto_post = f"⚠️ Demoras Administrativas (POST-Facturación) - Operaciones de: {asesor} - Sucursal: {sucursal}"
                                        saludo_post = f"Hola <b>Equipo de Administración ({sucursal})</b>,"
                                        intro_post = f"El sistema detectó demoras en Facturación o etapas posteriores correspondientes a operaciones del asesor <b>{asesor}</b> (en copia). Se solicita gestión administrativa inmediata para evitar mayores retrasos en la entrega. <b>Por favor, ingresen a la aplicación de Flujo de Operaciones y, en la pestaña de 'Alertas de Demora', completen la casilla de 'Comentario' para informar el estado actual de cada caso:</b>"
                                        
                                        exito, _ = enviar_correo_personalizado(df_asesor, destinatario=correos_admin_destinatario, cc=correo_asesor, asunto=asunto_post, saludo=saludo_post, intro_texto=intro_post)
                                        if exito: correos_enviados += 1
                                        
                                if correos_enviados > 0:
                                    st.success(f"¡Se enviaron exitosamente {correos_enviados} alertas escaladas con copia cruzada organizadas por sucursal!")
                                    if asesores_sin_correo:
                                        st.warning(f"Nota: Los siguientes asesores no estaban en el diccionario: {', '.join(set(asesores_sin_correo))}")
                            else:
                                st.error("No se encontró la columna de Asesores para dividir los correos.")
                                
            else:
                st.success("¡Excelente! No hay boletos demorados en este momento.")

    except Exception as e:
        st.error(f"Error procesando la base de datos. (Detalle: {e})")
    
