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
            
            columnas_mostrar = ['Identificador', 'Numero_de_Boleto__c', 'Sucursal_de_Venta__c', 'Denominacion_Comercial__c', 'Nombres_de_Cuenta__c', 'Estado Actual', 'Días en este estado', 'Fecha de Último Estado', 'Nombre_Asesor__c', 'Comentario']
            columnas_relevantes = [c for c in columnas_mostrar if c in df_tabla.columns]
            
            # AQUI ESTA EL CAMBIO: Pasamos el dataframe limpio, sin los colores (.style), para que Streamlit nos deje escribir.
            df_editado = st.data_editor(
                df_tabla[columnas_relevantes],
                use_container_width=True, hide_index=True,
                column_config={"Identificador": None, "Comentario": st.column_config.TextColumn("💬 Comentario (Doble clic)", help="Escribe el motivo de la demora.")},
                disabled=[c for c in columnas_relevantes if c != 'Comentario'], key="editor_tabla_auditoria"
            )
            
            if df_editado is not None:
                for idx in df_editado.index:
                    df.loc[df['Identificador'] == df_editado.at[idx, 'Identificador'], 'Comentario'] = df_editado.at[idx, 'Comentario']
                df[['Identificador', 'Comentario']].drop_duplicates(subset=['Identificador']).to_csv(ARCHIVO_COMENTARIOS, index=False)
                
                st.download_button("📥 Descargar tabla de Auditoría (Excel)", data=convertir_df_a_excel(df_editado), file_name=f'Auditoria_Boletos.xlsx', mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
