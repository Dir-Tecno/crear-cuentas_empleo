import os
import pandas as pd
import glob
from datetime import datetime

# ==================== CONFIGURACIÓN ====================

BASE_DIR = r"D:\TEMP"
#BASE_DIR = r"D:\DESARROLLO\ECONOMIA_SOCIAL\Empleo_Apertura_Cuenta_Bancaria"
FERIAS_NOC_DIR = os.path.join(BASE_DIR, "PPP")
PROCESADOS_DIR = os.path.join(FERIAS_NOC_DIR, "procesados_directo")

# ==================== FUNCIONES DE FORMATO ====================

def formatear_campo(valor, longitud, tipo, default=''):
    """
    Formatea un campo según el tipo y longitud especificados.
    
    Args:
        valor: Valor a formatear
        longitud: Longitud total del campo
        tipo: 'A' para alfanumérico, 'N' para numérico
        default: Valor por defecto si valor es None o vacío
    
    Returns:
        String formateado con la longitud correcta
    """
    # Si es una Serie de pandas, tomar el primer valor
    if isinstance(valor, pd.Series):
        valor = valor.iloc[0] if len(valor) > 0 else default
    
    # Usar default si el valor está vacío o es None
    if pd.isna(valor) or valor == '' or valor is None:
        valor = default
    
    # Convertir a string
    valor_str = str(valor).strip()
    
    if tipo == 'N':
        # Numérico: rellenar con ceros a la izquierda
        # Eliminar caracteres no numéricos
        valor_str = ''.join(c for c in valor_str if c.isdigit())
        if valor_str == '':
            valor_str = '0'
        return valor_str.zfill(longitud)[:longitud]
    else:
        # Alfanumérico: rellenar con espacios a la derecha
        return valor_str.ljust(longitud)[:longitud]


def procesar_apellido(apellido):
    """Procesa el apellido: si hay más de 1 apellido, colocar el 1ero y 2do si existiera."""
    if pd.isna(apellido) or apellido == '':
        return '', ''
    
    apellidos = str(apellido).strip().split()
    primer_apellido = apellidos[0] if len(apellidos) > 0 else ''
    segundo_apellido = apellidos[1] if len(apellidos) > 1 else ''
    
    return primer_apellido, segundo_apellido


def procesar_nombre(nombre):
    """Procesa el nombre: si hay más de 1 nombre, colocar el 1ero y 2do si existiera."""
    if pd.isna(nombre) or nombre == '':
        return '', ''
    
    nombres = str(nombre).strip().split()
    primer_nombre = nombres[0] if len(nombres) > 0 else ''
    segundo_nombre = nombres[1] if len(nombres) > 1 else ''
    
    return primer_nombre, segundo_nombre


def procesar_celular_post(celular_post):
    """Procesa CELULAR_POST para extraer prefijo y número."""
    if pd.isna(celular_post) or celular_post == '':
        return '', ''
    
    celular_str = str(celular_post).strip()
    
    # Quitar el primer caracter si es "0"
    if celular_str.startswith('0'):
        celular_str = celular_str[1:]
    
    # Verificar si los primeros 2 caracteres son "11" (Buenos Aires)
    if celular_str[:2] == '11':
        return celular_str[:2], celular_str[2:]
    # Verificar si los primeros 3 caracteres están en la lista
    elif celular_str[:3] in ['351', '358', '353']:
        return celular_str[:3], celular_str[3:]
    else:
        # Tomar los primeros 4 caracteres como prefijo
        return celular_str[:4], celular_str[4:]


def mapear_sexo(sexo):
    """
    Convierte el campo SEXO a formato HAB (1 o 2):
    - 'MUJER' o 'F' o '2' o '02' → '2'
    - 'VARON' o 'M' o '1' o '01' → '1'
    """
    if pd.isna(sexo):
        return ''
    
    sexo_str = str(sexo).strip().upper()
    
    # Mapear a formato HAB (1 = VARON, 2 = MUJER)
    if sexo_str in ['MUJER', 'F', '2', '02']:
        return '2'
    elif sexo_str in ['VARON', 'M', '1', '01']:
        return '1'
    else:
        return ''


def mapear_sexo_hab(id_sexo):
    """
    Convierte ID_SEXO a formato HAB:
    - '1' o '01' → 'M'
    - '2' o '02' → 'F'
    """
    if pd.isna(id_sexo):
        return ''
    
    sexo_str = str(id_sexo).strip()
    if sexo_str == '01' or sexo_str == '1':
        return 'M'
    elif sexo_str == '02' or sexo_str == '2':
        return 'F'
    else:
        return ''


def sanitizar_texto(texto):
    """
    Sanitiza un texto eliminando acentos y caracteres especiales.
    """
    if pd.isna(texto) or texto == '':
        return texto
    
    texto_str = str(texto)
    
    # Reemplazar vocales con acentos
    reemplazos = {
        'á': 'a', 'é': 'e', 'í': 'i', 'ó': 'o', 'ú': 'u',
        'Á': 'A', 'É': 'E', 'Í': 'I', 'Ó': 'O', 'Ú': 'U',
        'ä': 'a', 'ë': 'e', 'ï': 'i', 'ö': 'o', 'ü': 'u',
        'Ä': 'A', 'Ë': 'E', 'Ï': 'I', 'Ö': 'O', 'Ü': 'U',
        'à': 'a', 'è': 'e', 'ì': 'i', 'ò': 'o', 'ù': 'u',
        'À': 'A', 'È': 'E', 'Ì': 'I', 'Ò': 'O', 'Ù': 'U'
    }
    
    for acento, sin_acento in reemplazos.items():
        texto_str = texto_str.replace(acento, sin_acento)
    
    # Eliminar apóstrofes
    texto_str = texto_str.replace("'", "")
    
    return texto_str


# ==================== MAPEO DE CAMPOS ====================

def aplicar_logica_apoderado(row):
    """
    Aplica la lógica de apoderado según las reglas:
    - Si IdApoderado no es null/vacío: usar campos del apoderado
    - Si no: usar SEXO, NUMERO_DOCUMENTO, APELLIDO, etc.
    
    Returns:
        Dict con los campos procesados (SEXO, NRO_DOCUMENTO, APELLIDO, NOMBRE, etc.)
    """
    datos_procesados = {}
    
    # Verificar si tiene apoderado: validar que IdApoderado no sea null/vacío
    tiene_apoderado = False
    if 'IdApoderado' in row.index:
        IdApoderado = row.get('IdApoderado', '')
        if not pd.isna(IdApoderado) and str(IdApoderado).strip() != '':
            tiene_apoderado = True
    
    if tiene_apoderado:
        # Usar campos de apoderado
        datos_procesados['SEXO'] = row.get('APO_SEXO', '')
        datos_procesados['NRO_DOCUMENTO'] = row.get('APO_DNI', '')
        datos_procesados['APELLIDO'] = sanitizar_texto(row.get('APO_APELLIDO', ''))
        datos_procesados['NOMBRE'] = sanitizar_texto(row.get('APO_NOMBRE', ''))
        datos_procesados['CUIL'] = row.get('APO_CUIL', '')
        datos_procesados['FEC_NACIMIENTO'] = row.get('APO_FEC_NAC', '')
        datos_procesados['CELULAR_POST'] = row.get('APO_CELULAR', '')
        datos_procesados['MAIL_POST'] = row.get('APO_EMAIL', '')
        datos_procesados['N_CALLE'] = row.get('APO_CALLE', '')
        datos_procesados['ALTURA_ORACLE'] = row.get('APO_NRO', '')
        datos_procesados['N_BARRIO'] = row.get('APO_BARRIO', '')
        datos_procesados['N_LOCALIDAD'] = row.get('APO_LOCALIDAD', '')
        datos_procesados['CPA'] = row.get('APO_CP', '')
        datos_procesados['COD_BCO_CBA'] = row.get('APO_COD_SUC', '')
    else:
        # Usar campos normales del beneficiario
        datos_procesados['SEXO'] = row.get('SEXO', '')
        datos_procesados['NRO_DOCUMENTO'] = row.get('NUMERO_DOCUMENTO', '')
        datos_procesados['APELLIDO'] = sanitizar_texto(row.get('APELLIDO', ''))
        datos_procesados['NOMBRE'] = sanitizar_texto(row.get('NOMBRE', ''))
        datos_procesados['CUIL'] = row.get('CUIL', '')
        datos_procesados['FEC_NACIMIENTO'] = row.get('FER_NAC', '')
        datos_procesados['CELULAR_POST'] = row.get('TEL_CELULAR', '')
        datos_procesados['MAIL_POST'] = row.get('MAIL', '')
        datos_procesados['N_CALLE'] = row.get('CALLE', '')
        datos_procesados['ALTURA_ORACLE'] = row.get('NUMERO', '')
        datos_procesados['N_BARRIO'] = row.get('BARRIO', '')
        datos_procesados['N_LOCALIDAD'] = row.get('N_LOCALIDAD', '')
        datos_procesados['CPA'] = row.get('CODIGO_POSTAL', '')
        datos_procesados['COD_BCO_CBA'] = row.get('BEN_COD_SUC', '')
    
    return datos_procesados


# ==================== GENERACIÓN DE ARCHIVO .HAB ====================

def generar_linea_hab(row):
    """Genera una línea del archivo .HAB según el formato especificado."""
    linea = ''
    
    # Aplicar lógica de apoderado
    datos = aplicar_logica_apoderado(row)
    
    # Procesar apellidos y nombres
    primer_apellido, segundo_apellido = procesar_apellido(datos.get('APELLIDO', ''))
    primer_nombre, segundo_nombre = procesar_nombre(datos.get('NOMBRE', ''))
    
    # Procesar CELULAR_POST
    pref_tel, tel_particular = procesar_celular_post(datos.get('CELULAR_POST', ''))
    
    # Procesar N_BARRIO: si es NULL usar "OTRO"
    n_barrio = datos.get('N_BARRIO', '')
    if pd.isna(n_barrio) or n_barrio == '' or n_barrio is None:
        n_barrio = 'OTRO'
    
    # Procesar EMAIL: si supera 30 caracteres usar email genérico
    mail_post = datos.get('MAIL_POST', '')
    if not pd.isna(mail_post) and mail_post != '' and len(str(mail_post)) > 30:
        mail_post = 'mailgenerica@bancor.com.ar'
    
    # Obtener fecha actual en formato YYYYMMDD
    fecha_hoy = datetime.now().strftime('%Y%m%d')
    
    # Construir cada campo según especificación
    linea += formatear_campo('A', 1, 'A')  # TIPO DE REGISTRO
    linea += formatear_campo(datos.get('COD_BCO_CBA', ''), 5, 'N')  # SUCURSAL
    linea += formatear_campo('1', 2, 'N', '1')  # MONEDA
    linea += formatear_campo('1', 3, 'N', '1')  # TIPO DOCUMENTO
    linea += formatear_campo(datos.get('NRO_DOCUMENTO', ''), 11, 'N')  # NRO DOCUMENTO
    linea += formatear_campo('7', 3, 'N')  # CLAVE FISCAL
    linea += formatear_campo(datos.get('CUIL', ''), 11, 'N')  # NRO CLAVE FISCAL
    linea += formatear_campo('0', 2, 'N', '0')  # TIPO CUENTA
    linea += formatear_campo('0', 9, 'N', '0')  # NRO CUENTA
    linea += formatear_campo(fecha_hoy, 8, 'N')  # FECHA ALTA (SYSDATE)
    linea += formatear_campo(primer_apellido, 15, 'A')  # PRIMER APELLIDO
    linea += formatear_campo(segundo_apellido, 15, 'A')  # SEGUNDO APELLIDO
    linea += formatear_campo(primer_nombre, 15, 'A')  # PRIMER NOMBRE
    linea += formatear_campo(segundo_nombre, 15, 'A')  # SEGUNDO NOMBRE
    linea += formatear_campo('4', 2, 'N', '4')  # CONDICION IVA
    linea += formatear_campo(datos.get('N_CALLE', ''), 30, 'A')  # DOMICILIO PARTICULAR
    linea += formatear_campo(datos.get('ALTURA_ORACLE', ''), 5, 'N')  # NRO DOMICILIO
    linea += formatear_campo('', 2, 'N')  # PISO
    linea += formatear_campo('', 3, 'N')  # DEPARTAMENTO
    linea += formatear_campo(n_barrio, 30, 'A')  # BARRIO
    linea += formatear_campo(datos.get('N_LOCALIDAD', ''), 30, 'A')  # LOCALIDAD
    linea += formatear_campo('4', 3, 'N', '4')  # CODIGO PROVINCIA
    linea += formatear_campo(datos.get('CPA', ''), 5, 'N')  # CODIGO POSTAL
    linea += formatear_campo('0', 8, 'N', '0')  # CODIGO POSTAL EXTENDIDO
    linea += formatear_campo(pref_tel, 5, 'A')  # PREF. TEL PARTICULAR
    linea += formatear_campo(tel_particular, 11, 'N')  # TEL PARTICULAR
    linea += formatear_campo(pref_tel, 5, 'A')  # PREF. TEL CEL
    linea += formatear_campo(tel_particular, 11, 'N')  # TEL MOVIL
    linea += formatear_campo(datos.get('N_CALLE', ''), 30, 'A')  # DOMICILIO COMERCIAL
    linea += formatear_campo(datos.get('ALTURA_ORACLE', ''), 5, 'N')  # NRO DOMICILIO COMERCIAL
    linea += formatear_campo('', 2, 'N')  # PISO COMERCIAL
    linea += formatear_campo('', 3, 'N')  # DEPARTAMENTO COMERCIAL
    linea += formatear_campo(n_barrio, 30, 'A')  # BARRIO COMERCIAL
    linea += formatear_campo(datos.get('N_LOCALIDAD', ''), 30, 'A')  # LOCALIDAD COMERCIAL
    linea += formatear_campo('4', 3, 'N', '4')  # COD. PROV. COMERC
    linea += formatear_campo(datos.get('CPA', ''), 5, 'N')  # COD POSTAL COMERCIAL
    linea += formatear_campo('0', 8, 'N', '0')  # COD POSTAL EXTENDIDO COMERC
    linea += formatear_campo(pref_tel, 5, 'A')  # PREF. TEL COMERCIAL
    linea += formatear_campo(tel_particular, 11, 'N')  # TELEFONO COMERC
    linea += formatear_campo(datos.get('FEC_NACIMIENTO', ''), 8, 'N')  # FECHA NACIMIENTO
    linea += formatear_campo('1', 4, 'N','1')  # ESTADO CIVIL
    linea += formatear_campo('S', 1, 'A', 'S')  # RESIDENTE
    linea += formatear_campo(mapear_sexo(datos.get('SEXO', '')), 1, 'A')  # SEXO
    linea += formatear_campo('1', 3, 'N', '1')  # NACIONALIDAD
    linea += formatear_campo(mail_post, 30, 'A')  # EMAIL
    linea += formatear_campo('F', 1, 'A', 'F')  # TIPO PERSONA
    linea += formatear_campo('2', 5, 'N', '2')  # COD. ACT. BCRA
    linea += formatear_campo('27', 3, 'N', '27')  # COD. NAT. JURIDICA
    linea += formatear_campo('', 15, 'A')  # PRIMER APELLIDO CONYUGE
    linea += formatear_campo('', 15, 'A')  # SEGUNDO APELLIDO CONYUGE
    linea += formatear_campo('', 15, 'A')  # PRIMER NOMBRE CONYUGE
    linea += formatear_campo('', 15, 'A')  # SEGUNDO NOMBRE CONYUGE
    linea += formatear_campo('', 1, 'A')  # SEXO CONYUGE
    linea += formatear_campo('', 3, 'N')  # TIPO DOC CONYUGE
    linea += formatear_campo('', 11, 'N')  # NRO DOC CONYUGE
    linea += formatear_campo('', 11, 'N')  # CUIT CONYUGE
    linea += formatear_campo('', 8, 'N')  # FECHA NACIMIENTO CONYUGE
    linea += formatear_campo('', 3, 'N')  # NACIONALIDAD CONYUGE
    linea += formatear_campo('1137', 5, 'N', '1137')  # NRO EMPRESA
    linea += formatear_campo('0', 3, 'N', '0')  # TIPO CONVENIO
    linea += formatear_campo('1', 1, 'A', '1')  # VALIDA NOMBRE
    linea += formatear_campo('', 30, 'A')  # NOMBRE CLIENTE SEGUN PATRON
    linea += formatear_campo('', 409, 'A')  # FILLER
    linea += formatear_campo('', 2, 'N')  # TIPO SOLICITUD
    linea += formatear_campo('', 22, 'A')  # CRU
    linea += formatear_campo('', 330, 'A')  # FILLER 2
    linea += formatear_campo('', 21, 'A')  # DATOS PARA EMPRESA
    # MOTIVO DEL ERROR 1-7 (7 campos de 5 caracteres cada uno)
    for _ in range(7):
        linea += formatear_campo('', 5, 'A')
    
    return linea


def generar_archivo_hab(df: pd.DataFrame, output_path: str) -> tuple:
    """
    Genera un archivo .HAB a partir de un DataFrame.
    
    Validación: Solo genera línea HAB si IdApoderado NO es null/vacío.
    
    Args:
        df: DataFrame con los datos procesados
        output_path: Ruta donde guardar el archivo .HAB
    
    Returns:
        Tupla (lineas_generadas, lineas_saltadas) - número de líneas generadas y saltadas
    """
    lineas_generadas = 0
    lineas_saltadas = 0
    
    # newline='' para controlar manualmente los saltos de línea (CR-LF para Windows)
    with open(output_path, 'w', encoding='latin-1', newline='') as f:
        for _, row in df.iterrows():
            # Validación: Si IdApoderado está vacío o es null, SALTAR registro
            IdApoderado = row.get('IdApoderado', '')
            if pd.isna(IdApoderado) or str(IdApoderado).strip() == '':
                lineas_saltadas += 1
                continue
            
            linea = generar_linea_hab(row)
            f.write(linea + '\r\n')  # CR-LF (Windows)
            lineas_generadas += 1
    
    return lineas_generadas, lineas_saltadas


# ==================== PROCESAMIENTO PRINCIPAL ====================

def procesar_archivo_excel(excel_path):
    """Procesa un archivo Excel individual y genera el archivo .HAB"""
    filename = os.path.basename(excel_path)
    print(f"\n🔄 Procesando archivo: {filename}")
    
    try:
        # Leer archivo Excel
        df = pd.read_excel(excel_path, dtype=str)
        
        # Limpiar nombres de columnas (quitar espacios al inicio/final)
        df.columns = df.columns.str.strip()
        
        print(f"   ✅ Archivo cargado con {len(df)} filas y {len(df.columns)} columnas.")
        print(f"   📋 Columnas disponibles: {list(df.columns)}")
        
        # Validar columnas mínimas requeridas
        # Si tiene apoderado, verificar campos de apoderado
        # Si no, verificar campos de beneficiario
        tiene_columnas_apoderado = all(col in df.columns for col in ['IdApoderado', 'APO_SEXO'])
        tiene_columnas_beneficiario = all(col in df.columns for col in ['NUMERO_DOCUMENTO', 'SEXO'])
        
        if not tiene_columnas_apoderado and not tiene_columnas_beneficiario:
            print(f"   ❌ Error: El archivo debe contener campos de beneficiario o apoderado")
            print(f"   💡 Campos mínimos beneficiario: SEXO, NUMERO_DOCUMENTO, APELLIDO, NOMBRE, CUIL")
            print(f"   💡 Campos mínimos apoderado: APO_SEXO, IdApoderado, APO_APELLIDO, APO_NOMBRE, APO_CUIL")
            return
        
        # Contar registros con apoderado válido (IdApoderado no vacío)
        registros_con_apoderado = 0
        if 'IdApoderado' in df.columns:
            # Verificar que IdApoderado no esté vacío
            mask_apoderado = (df['IdApoderado'].notna()) & \
                             (df['IdApoderado'].astype(str).str.strip() != '')
            registros_con_apoderado = mask_apoderado.sum()
            print(f"   📊 Registros con apoderado válido (IdApoderado): {registros_con_apoderado}")
            print(f"   📊 Registros sin apoderado válido: {len(df) - registros_con_apoderado}")
        
        # Generar archivo .HAB
        print(f"   📝 Generando archivo .HAB...")
        
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        hab_filename = f"{os.path.splitext(filename)[0]}_{timestamp}.HAB"
        hab_path = os.path.join(PROCESADOS_DIR, hab_filename)
        
        lineas_hab, lineas_saltadas = generar_archivo_hab(df, hab_path)
        print(f"   ✅ Archivo .HAB generado: {hab_path}")
        print(f"   📊 Total de líneas creadas en archivo .HAB: {lineas_hab}")
        if lineas_saltadas > 0:
            print(f"   ⚠️  Registros saltados (IdApoderado vacío): {lineas_saltadas}")
        
        # Guardar también Excel procesado con los datos normalizados
        excel_output_filename = f"procesado_{os.path.splitext(filename)[0]}_{timestamp}.xlsx"
        excel_output_path = os.path.join(PROCESADOS_DIR, excel_output_filename)
        df.to_excel(excel_output_path, index=False)
        print(f"   💾 Excel procesado guardado: {excel_output_path}")
        
    except Exception as e:
        print(f"   ❌ Error procesando archivo {filename}: {e}")
        import traceback
        traceback.print_exc()


def main():
    """Función principal que procesa todos los archivos .xlsx"""
    print("🚀 Iniciando procesamiento directo de archivos Excel...")
    print(f"📁 Directorio de entrada: {FERIAS_NOC_DIR}")
    print(f"📁 Directorio de salida: {PROCESADOS_DIR}")
    
    # Crear directorio de salida si no existe
    os.makedirs(PROCESADOS_DIR, exist_ok=True)
    
    # Buscar archivos .xlsx en el directorio
    pattern = os.path.join(FERIAS_NOC_DIR, "*.xlsx")
    all_excel_files = glob.glob(pattern)
    
    # Filtrar archivos temporales de Excel
    excel_files = [f for f in all_excel_files if not os.path.basename(f).startswith('~$')]
    
    if not excel_files:
        print(f"❌ No se encontraron archivos .xlsx en {FERIAS_NOC_DIR}")
        print(f"\n💡 Uso alternativo: Puedes llamar a la función directamente:")
        print(f"   from procesar_excel_directo import procesar_archivo_excel")
        print(f"   procesar_archivo_excel('ruta/al/archivo.xlsx')")
        return
    
    print(f"📋 Se encontraron {len(excel_files)} archivo(s) .xlsx para procesar\n")
    
    archivos_procesados = 0
    archivos_con_error = 0
    
    for excel_file in excel_files:
        try:
            procesar_archivo_excel(excel_file)
            archivos_procesados += 1
        except Exception as e:
            print(f"❌ Error general procesando {os.path.basename(excel_file)}: {e}")
            archivos_con_error += 1
    
    print("\n🎉 Procesamiento completado!")
    print(f"📊 Archivos procesados exitosamente: {archivos_procesados}")
    if archivos_con_error > 0:
        print(f"⚠️  Archivos con error: {archivos_con_error}")


if __name__ == "__main__":
    main()
