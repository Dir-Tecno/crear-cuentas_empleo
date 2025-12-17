import streamlit as st
import pandas as pd
import io
from datetime import datetime
from procesar_excel_directo import (
    formatear_campo,
    procesar_apellido,
    procesar_nombre,
    procesar_celular_post,
    mapear_sexo,
    mapear_sexo_hab,
    sanitizar_texto,
    aplicar_logica_apoderado,
    generar_linea_hab,
    generar_archivo_hab
)

# Configuración de la página
st.set_page_config(
    page_title="Generador de Archivos HAB",
    page_icon="🏦",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ==================== SIDEBAR CON INSTRUCCIONES ====================
with st.sidebar:
    st.title("📋 Instrucciones")
    
    st.markdown("""
    ### ¿Qué hace esta aplicación?
    
    Convierte archivos Excel con datos de beneficiarios en archivos **.HAB** 
    (formato de ancho fijo) para el Banco de Córdoba.
    
    ---
    
    ### 📊 Formato del Excel esperado
    
    #### **Campos del Beneficiario (obligatorios si no hay apoderado):**
    - `SEXO` - Género (VARON/MUJER)
    - `NUMERO_DOCUMENTO` - DNI
    - `APELLIDO` - Apellidos completos
    - `NOMBRE` - Nombres completos
    - `CUIL` - CUIL del beneficiario
    - `FER_NAC` - Fecha de nacimiento (YYYYMMDD)
    - `TEL_CELULAR` - Número de celular
    - `MAIL` - Email (máx. 30 caracteres)
    - `CALLE` - Nombre de calle
    - `NUMERO` - Altura
    - `BARRIO` - Barrio (si es NULL usa "OTRO")
    - `N_LOCALIDAD` - Localidad
    - `CODIGO_POSTAL` - CP
    - `BEN_COD_SUC` - Código de sucursal
    
    #### **Campos del Apoderado :**
    - `ID_APODERADO` -ID del apoderado 
    - `APO_DNI` - DNI del apoderado 
    - `APO_SEXO` - Género del apoderado
    - `APO_APELLIDO` - Apellidos del apoderado
    - `APO_NOMBRE` - Nombres del apoderado
    - `APO_CUIL` - CUIL del apoderado
    - `APO_FEC_NAC` - Fecha de nacimiento
    - `APO_CELULAR` - Celular del apoderado
    - `APO_EMAIL` - Email del apoderado
    - `APO_CALLE` - Calle del apoderado
    - `APO_NRO` - Altura del apoderado
    - `APO_BARRIO` - Barrio del apoderado
    - `APO_LOCALIDAD` - Localidad del apoderado
    - `APO_CP` - CP del apoderado
    - `APO_COD_SUC` - Código de sucursal
    
    ---
    
    ### ⚙️ Procesamiento automático
    
    - ✅ **Limpieza de columnas:** Elimina espacios en nombres de columnas
    - ✅ **Sanitización:** Elimina acentos y caracteres especiales
    - ✅ **Nombres/Apellidos:** Separa primer y segundo nombre/apellido
    - ✅ **Celulares:** Extrae prefijo (11, 351, 358, 353, etc.) y número
    - ✅ **Emails largos:** Si supera 30 caracteres → email genérico
    - ✅ **Barrios vacíos:** Si es NULL → "OTRO"
    - ✅ **Apoderados:** Si TIENE_APODERADO='S' Y APO_DNI tiene valor → usa datos APO_*
    - ✅ **Mapeo SEXO:** MUJER/VARON → 2/1 para formato HAB
    
    ---
    
    ### 🎯 Pasos de uso
    
    1. **Cargar** el archivo Excel (.xlsx)
    2. **Verificar** la vista previa y columnas detectadas
    3. **Revisar** registros con/sin apoderado (si aplica)
    4. **Generar** y **descargar** el archivo .HAB
    
    ---
    
    ### 📝 Formato del archivo .HAB
    
    - **Encoding:** latin-1 (compatibilidad bancaria)
    - **Saltos de línea:** CR-LF (Windows)
    - **Campo SEXO:** 1=VARON, 2=MUJER
    - **Ancho fijo:** Cada campo tiene longitud específica
    """)
    
    st.markdown("---")
    st.success("✨ **¡Listo para usar!** La app procesa automáticamente todos los formatos.")

# ==================== CONTENIDO PRINCIPAL ====================

st.title("🏦 Generador de Archivos HAB para Banco de Córdoba")
st.markdown("### Convierte datos de Excel a formato HAB")

# Sección de carga de archivo
st.markdown("---")
uploaded_file = st.file_uploader(
    "📁 Selecciona un archivo Excel (.xlsx)",
    type=['xlsx'],
    help="Carga un archivo Excel con los campos requeridos (ver sidebar)"
)

if uploaded_file is not None:
    try:
        # Leer el archivo Excel
        df = pd.read_excel(uploaded_file, dtype=str)
        
        # Limpiar nombres de columnas (quitar espacios al inicio/final)
        df.columns = df.columns.str.strip()
        
        # Mostrar información del archivo
        st.success(f"✅ Archivo cargado exitosamente: **{uploaded_file.name}**")
        
        col1, col2 = st.columns(2)
        with col1:
            st.metric("Total de registros", len(df))
        with col2:
            st.metric("Total de columnas", len(df.columns))
        
        # Validar columnas
        tiene_columnas_apoderado = all(col in df.columns for col in ['ID_APODERADO', 'APO_SEXO'])
        tiene_columnas_beneficiario = all(col in df.columns for col in ['NUMERO_DOCUMENTO', 'SEXO'])
        
        if not tiene_columnas_apoderado and not tiene_columnas_beneficiario:
            st.error("❌ Error: El archivo debe contener campos de beneficiario o apoderado")
            st.info("💡 **Campos mínimos beneficiario:** SEXO, NUMERO_DOCUMENTO, APELLIDO, NOMBRE, CUIL")
            st.info("💡 **Campos mínimos apoderado:** APO_SEXO, ID_APODERADO, APO_APELLIDO, APO_NOMBRE, APO_CUIL")
        else:
            # Contar registros con apoderado válido
            registros_con_apoderado = 0
            if 'ID_APODERADO' in df.columns:
                # Verificar que ID_APODERADO no esté vacío
                mask_apoderado = (df['ID_APODERADO'].notna()) & \
                                 (df['ID_APODERADO'].astype(str).str.strip() != '')
                registros_con_apoderado = mask_apoderado.sum()
                
                col1, col2 = st.columns(2)
                with col1:
                    st.info(f"👤 Registros con apoderado válido: **{registros_con_apoderado}**")
                with col2:
                    st.info(f"👤 Registros sin apoderado válido: **{len(df) - registros_con_apoderado}**")
            
            # Vista previa de datos
            with st.expander("👁️ Ver vista previa de los datos (primeras 10 filas)"):
                st.dataframe(df.head(10), use_container_width=True)
            
            # Mostrar columnas disponibles
            with st.expander("📋 Ver columnas disponibles"):
                st.write(list(df.columns))
            
            # DEBUG: Mostrar primer registro completo
            with st.expander("🔍 DEBUG - Ver primer registro completo"):
                if len(df) > 0:
                    primer_registro = df.iloc[0]
                    st.markdown("**Todas las columnas y valores del primer registro:**")
                    for col in df.columns:
                        valor = primer_registro[col]
                        st.text(f"{col}: '{valor}'")
                else:
                    st.warning("No hay registros en el archivo")
            
            st.markdown("---")
            
            # Opción para filtrar por apoderado
            st.markdown("### ⚙️ Opciones de procesamiento")
            solo_con_apoderado = st.checkbox(
                "✅ Procesar solo registros con ID_APODERADO válido",
                value=False,
                help="Si está marcado, solo se procesarán registros donde ID_APODERADO tenga valor. Si no, se procesan todos los registros."
            )
            
            st.markdown("---")
            
            # Botón para generar archivo HAB
            if st.button("🚀 Generar archivo .HAB", type="primary", use_container_width=True):
                with st.spinner("Procesando archivo..."):
                    try:
                        # Generar archivo HAB en memoria
                        output = io.StringIO(newline='')
                        lineas_generadas = 0
                        lineas_saltadas = 0
                        
                        for _, row in df.iterrows():
                            # Validación condicional según checkbox
                            if solo_con_apoderado:
                                ID_APODERADO = row.get('ID_APODERADO', '')
                                if pd.isna(ID_APODERADO) or str(ID_APODERADO).strip() == '':
                                    lineas_saltadas += 1
                                    continue
                            
                            linea = generar_linea_hab(row)
                            output.write(linea + '\r\n')  # CR-LF (Windows)
                            lineas_generadas += 1
                        
                        # Obtener contenido del archivo
                        hab_content = output.getvalue()
                        output.close()
                        
                        # Convertir a bytes con encoding latin-1
                        hab_bytes = hab_content.encode('latin-1')
                        
                        # Generar nombre de archivo con timestamp
                        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                        original_name = uploaded_file.name.rsplit('.', 1)[0]
                        hab_filename = f"{original_name}_{timestamp}.HAB"
                        
                        st.success(f"✅ Archivo .HAB generado exitosamente!")
                        
                        col1, col2 = st.columns(2)
                        with col1:
                            st.info(f"📊 Líneas creadas: **{lineas_generadas}**")
                        with col2:
                            if lineas_saltadas > 0 and solo_con_apoderado:
                                st.warning(f"⚠️  Registros saltados (ID_APODERADO vacío): **{lineas_saltadas}**")
                        
                        # Botón de descarga
                        st.download_button(
                            label="⬇️ Descargar archivo .HAB",
                            data=hab_bytes,
                            file_name=hab_filename,
                            mime="text/plain",
                            use_container_width=True
                        )
                        
                        st.balloons()
                        
                    except Exception as e:
                        st.error(f"❌ Error al generar el archivo .HAB: {str(e)}")
                        with st.expander("Ver detalles del error"):
                            import traceback
                            st.code(traceback.format_exc())
    
    except Exception as e:
        st.error(f"❌ Error al leer el archivo Excel: {str(e)}")
        with st.expander("Ver detalles del error"):
            import traceback
            st.code(traceback.format_exc())

else:
    # Mensaje de bienvenida cuando no hay archivo cargado
    st.info("👆 Por favor, carga un archivo Excel para comenzar")

# Footer
st.markdown("---")
st.markdown(
    """
    <div style='text-align: center; color: #666;'>
        <p>Generador de Archivos HAB v1.0 | Banco de Córdoba</p>
    </div>
    """,
    unsafe_allow_html=True
)
