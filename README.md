# 🏦 Generador de Archivos HAB - Banco de Córdoba

Aplicación web para convertir archivos Excel con datos de beneficiarios en archivos .HAB (formato de ancho fijo) para el Banco de Córdoba.

## 📋 Descripción

Esta aplicación procesa archivos Excel que contienen información de beneficiarios y/o apoderados, generando archivos .HAB con el formato requerido por el banco para la apertura de cuentas.

## 🚀 Inicio Rápido

### Instalación

1. **Clonar el repositorio** (si aplica)
```bash
cd crear-cuentas_empleo
```

2. **Crear un entorno virtual** (recomendado)
```bash
python -m venv venv
```

3. **Activar el entorno virtual**
   - Windows:
   ```bash
   venv\Scripts\activate
   ```
   - Linux/Mac:
   ```bash
   source venv/bin/activate
   ```

4. **Instalar dependencias**
```bash
pip install -r requirements.txt
```

### Ejecutar la aplicación

```bash
streamlit run app.py
```

La aplicación se abrirá automáticamente en tu navegador en `http://localhost:8501`

## 📊 Uso de la Aplicación

1. **Cargar el archivo Excel** con los datos de beneficiarios
2. **Verificar** la vista previa de los datos cargados
3. **Generar** el archivo .HAB presionando el botón
4. **Descargar** el archivo .HAB generado

## 📁 Formato del Excel

### Campos del Beneficiario (obligatorios si no hay apoderado):
- `SEXO` - Género (MUJER/VARON)
- `NUMERO_DOCUMENTO` - DNI
- `APELLIDO` - Apellidos completos
- `NOMBRE` - Nombres completos
- `CUIL` - CUIL del beneficiario
- `FER_NAC` - Fecha de nacimiento (YYYYMMDD)
- `TEL_CELULAR` - Número de celular
- `MAIL` - Email (máx. 30 caracteres)
- `CALLE` - Nombre de calle
- `NUMERO` - Altura
- `BARRIO` - Barrio
- `N_LOCALIDAD` - Localidad
- `CODIGO_POSTAL` - CP
- `BEN_COD_SUC` - Código de sucursal

### Campos del Apoderado (opcionales):
- `TIENE_APODERADO` - Indica si tiene apoderado
- `APO_DNI` - DNI del apoderado
- `APO_SEXO` - Género
- `APO_APELLIDO`, `APO_NOMBRE`, `APO_CUIL`, etc.

Ver el sidebar de la aplicación para la lista completa de campos.

## ⚙️ Procesamiento Automático

La aplicación realiza automáticamente:
- ✅ **Limpieza de columnas**: Elimina espacios al inicio/final de nombres de columnas
- ✅ **Sanitización de texto**: Elimina acentos y caracteres especiales
- ✅ **Separación de nombres**: Divide primer y segundo nombre/apellido
- ✅ **Procesamiento de celulares**: Extrae prefijo y número según código de área
- ✅ **Validación de emails**: Reemplaza emails largos (>30 caracteres) con genérico
- ✅ **Lógica de apoderado**: Si TIENE_APODERADO='S' y APO_DNI tiene valor, usa datos del apoderado
- ✅ **Mapeo de SEXO**: Convierte 'MUJER'/'VARON' a '2'/'1' para formato HAB
- ✅ **Formato Windows**: Genera archivo con saltos de línea CR-LF
- ✅ **Encoding bancario**: Usa latin-1 para compatibilidad

## 🛠️ Uso del Script Original (CLI)

También puedes usar el script original desde línea de comandos:

```python
from procesar_excel_directo import procesar_archivo_excel
procesar_archivo_excel('ruta/al/archivo.xlsx')
```

## 📦 Archivos del Proyecto

- `app.py` - Aplicación Streamlit (interfaz web)
- `procesar_excel_directo.py` - Lógica de procesamiento
- `requirements.txt` - Dependencias del proyecto
- `README.md` - Este archivo

## 🔧 Tecnologías

- Python 3.x
- Streamlit - Framework de aplicación web
- Pandas - Procesamiento de datos
- OpenPyXL - Lectura de archivos Excel

## 📝 Notas Importantes

### Formato del archivo HAB:
- **Encoding**: latin-1 (compatibilidad bancaria)
- **Saltos de línea**: CR-LF (formato Windows)
- **Campo SEXO**: '1' = VARON, '2' = MUJER

### Lógica de apoderado:
- Se requiere `TIENE_APODERADO = 'S'` **Y** que `APO_DNI` tenga valor
- Cuando hay apoderado válido, se usan **todos** los datos del apoderado (APO_*)
- Los campos del beneficiario se usan solo cuando NO hay apoderado

### Procesamiento de campos:
- **Nombres de columnas**: Se limpian espacios automáticamente
- **Emails largos**: Si supera 30 caracteres → `mailgenerica@bancor.com.ar`
- **Barrios vacíos**: Si es NULL → 'OTRO'
- **Acentos**: Se eliminan automáticamente de apellidos y nombres
