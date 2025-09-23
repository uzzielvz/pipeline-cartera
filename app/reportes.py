from flask import Blueprint, render_template, request, send_file, flash, redirect, url_for, Response
import pandas as pd
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter
from openpyxl.styles import numbers
import os
import re
import logging
from datetime import datetime
from werkzeug.utils import secure_filename
import urllib.parse
from config import (
    ALLOWED_EXTENSIONS, UPLOAD_FOLDER, MAX_FILE_SIZE, COLUMN_MAPPING, 
    DTYPE_CONFIG, LISTA_FRAUDE, EXCEL_CONFIG, COLORS, ADDITIONAL_COLUMNS,
    MORA_BLUE_COLUMNS, CURRENCY_COLUMNS_KEYWORDS, DATE_COLUMNS_KEYWORDS
)

# Configurar logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

reportes_bp = Blueprint('reportes', __name__)

def allowed_file(filename):
    """Verifica que el archivo sea Excel"""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def validate_file_size(file_path):
    """Valida que el archivo no exceda el tamaño máximo permitido"""
    try:
        file_size = os.path.getsize(file_path)
        if file_size > MAX_FILE_SIZE:
            raise ValueError(f"El archivo excede el tamaño máximo de {MAX_FILE_SIZE // (1024*1024)}MB")
        return True
    except OSError as e:
        raise ValueError(f"Error al verificar el tamaño del archivo: {str(e)}")

def clean_dataframe_columns(df):
    """Limpia los nombres de columnas del DataFrame"""
    df.columns = df.columns.str.replace('\n', ' ').str.strip()
    return df

def standardize_codes(df, code_columns):
    """Estandariza códigos a 6 dígitos con ceros a la izquierda"""
    for col in code_columns:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0).astype(int).astype(str).str.zfill(6)
    return df

def clean_phone_numbers(df):
    """Limpia y estandariza números de teléfono"""
    for col in df.columns:
        if 'Teléfono' in col:
            df[col] = df[col].fillna('').astype(str)
    return df

def add_geolocation_links(df, geolocation_column):
    """Añade columnas de enlaces de geolocalización al DataFrame"""
    if geolocation_column in df.columns:
        geolocalizacion_data = df[geolocation_column].apply(generar_link_google_maps)
        df['link_texto'] = [item[0] for item in geolocalizacion_data]
        df['link_url'] = [item[1] for item in geolocalizacion_data]
    return df

def add_par_column(df, mora_column):
    """Añade columna PAR y la posiciona correctamente"""
    df['PAR'] = df[mora_column].apply(asignar_rango_mora)
    
    # Reordenar columnas para que 'PAR' esté al lado de 'Días de mora'
    columnas = df.columns.tolist()
    mora_index = columnas.index(mora_column)
    nuevas_columnas = (columnas[:mora_index+1] + 
                      ['PAR'] + 
                      columnas[mora_index+1:])
    
    return df[nuevas_columnas]

def create_excel_hyperlink(worksheet, row, col, text, url):
    """Crea un hipervínculo en Excel de forma segura"""
    try:
        if pd.notna(url) and str(url).strip() and str(url) != 'nan':
            worksheet.cell(row=row, column=col).hyperlink = str(url)
            worksheet.cell(row=row, column=col).value = str(text)
        else:
            worksheet.cell(row=row, column=col).value = str(text)
    except Exception as e:
        logger.warning(f"Error creando hipervínculo en fila {row}, columna {col}: {str(e)}")
        worksheet.cell(row=row, column=col).value = str(text)

def generate_valid_table_name(sheet_name):
    """
    Genera un nombre de tabla válido para Excel que siempre empiece con una letra.
    
    Reglas de Excel para nombres de tabla:
    - Debe empezar con una letra o guión bajo
    - No puede empezar con un número
    - Solo puede contener letras, números, guiones bajos y puntos
    - No puede contener espacios
    - Máximo 255 caracteres
    """
    # Limpiar el nombre de la hoja
    clean_name = str(sheet_name).replace(' ', '_').replace('-', '_')
    
    # Remover caracteres no válidos (mantener solo letras, números, guiones bajos y puntos)
    import re
    clean_name = re.sub(r'[^a-zA-Z0-9_.]', '', clean_name)
    
    # Asegurar que empiece con una letra o guión bajo
    if clean_name and clean_name[0].isdigit():
        clean_name = f"T_{clean_name}"
    elif not clean_name or not clean_name[0].isalpha():
        clean_name = f"T_{clean_name}" if clean_name else "T_Table"
    
    # Limitar longitud a 255 caracteres
    if len(clean_name) > 255:
        clean_name = clean_name[:255]
    
    return clean_name

def generar_link_google_maps(geolocalizacion):
    """
    Traductor de Direcciones - Convierte cualquier formato de geolocalización en enlaces de Google Maps.
    
    Casos manejados:
    - URL existente de Google Maps: ("Ver en mapa", url_original)
    - Coordenadas GPS: ("Ver en mapa", url_búsqueda)
    - Dirección de texto: ("Ver en mapa", url_búsqueda)
    - Vacío/nulo: ("Ver en mapa", url_generico)
    """
    # Caso 1: Valor vacío o nulo
    if pd.isna(geolocalizacion) or str(geolocalizacion).strip() == '':
        return ("Ver en mapa", "https://maps.google.com")
    
    geolocalizacion = str(geolocalizacion).strip()
    
    # Caso 2: Ya es una URL de Google Maps
    if "http" in geolocalizacion or "google.com/maps" in geolocalizacion:
        return ("Ver en mapa", geolocalizacion)
    
    # Caso 3: Contiene coordenadas GPS (patrón con °, ', ", N, W, S, E)
    if any(char in geolocalizacion for char in ['°', "'", '"', 'N', 'W', 'S', 'E']):
        try:
            # Extraer coordenadas usando regex
            # Patrón para coordenadas: 19°12'12.2"N 100°07'51.8"W
            coord_pattern = r"(\d+)°(\d+)'([\d.]+)\"([NS])\s+(\d+)°(\d+)'([\d.]+)\"([WE])"
            match = re.search(coord_pattern, geolocalizacion)
            
            if match:
                lat_deg, lat_min, lat_sec, lat_dir = match.groups()[:4]
                lon_deg, lon_min, lon_sec, lon_dir = match.groups()[4:]
                
                # Convertir a decimal
                lat_decimal = float(lat_deg) + float(lat_min)/60 + float(lat_sec)/3600
                lon_decimal = float(lon_deg) + float(lon_min)/60 + float(lon_sec)/3600
                
                # Aplicar dirección (N/S, E/W)
                if lat_dir == 'S':
                    lat_decimal = -lat_decimal
                if lon_dir == 'W':
                    lon_decimal = -lon_decimal
                
                # Crear URL de Google Maps
                url = f"https://www.google.com/maps/search/?api=1&query={lat_decimal},{lon_decimal}"
                return ("Ver en mapa", url)
        except:
            pass
    
    # Caso 4: Dirección de texto - crear búsqueda
    direccion_encoded = urllib.parse.quote_plus(geolocalizacion)
    url = f"https://www.google.com/maps/search/?api=1&query={direccion_encoded}"
    return ("Ver en mapa", url)

def asignar_rango_mora(dias_mora):
    """
    Asigna valor PAR (Período de Antigüedad de Recuperación) basado en días de mora.
    
    Reglas de categorización:
    - 1-7 días: '7'
    - 8-15 días: '15' 
    - 16-30 días: '30'
    - 31-60 días: '60'
    - 61-90 días: '90'
    - >90 días: '>90'
    - Otros casos: 'N/A'
    """
    if pd.isna(dias_mora) or dias_mora < 1:
        return 'N/A'
    elif 1 <= dias_mora <= 7:
        return '7'
    elif 8 <= dias_mora <= 15:
        return '15'
    elif 16 <= dias_mora <= 30:
        return '30'
    elif 31 <= dias_mora <= 60:
        return '60'
    elif 61 <= dias_mora <= 90:
        return '90'
    elif dias_mora > 90:
        return '>90'
    else:
        return 'N/A'

def escribir_hipervinculo_excel(worksheet, row, col, texto, url):
    """Escribe un hipervínculo en una celda de Excel"""
    cell = worksheet.cell(row=row, column=col)
    cell.value = texto
    cell.hyperlink = url
    cell.font = Font(color="0000FF", underline="single")

def aplicar_formato_final(worksheet, df, es_hoja_mora=False):
    """Autoajuste de columnas, formato de moneda, fecha corta, y formatos especiales."""
    from openpyxl.styles import Font, PatternFill, Alignment
    
    # a) Autoajuste de columnas
    for i in range(1, worksheet.max_column + 1):
        column_letter = get_column_letter(i)
        max_length = 0
        for j in range(1, worksheet.max_row + 1):
            value = worksheet.cell(row=j, column=i).value
            if value is not None:
                max_length = max(max_length, len(str(value)))
        worksheet.column_dimensions[column_letter].width = min(max_length + 2, 50)

    # b) Formato de encabezados (Fila 2)
    worksheet.row_dimensions[2].height = EXCEL_CONFIG['header_height']
    for cell in worksheet[2]:
        cell.font = Font(bold=True)
    
    # c) Relleno azul en "Días de mora" (en todas las hojas)
    for cell in worksheet[2]:
        if cell.value == 'Días de mora':
            cell.fill = PatternFill(start_color=COLORS['light_blue'], end_color=COLORS['light_blue'], fill_type="solid")
            break

    # d) Formato de moneda en columnas conocidas del df
    columnas_moneda = [
        col for col in df.columns
        if any(key in col.lower() for key in CURRENCY_COLUMNS_KEYWORDS)
    ]
    for col_name in columnas_moneda:
        if col_name in df.columns:
            # Buscar índice de la columna por encabezado (fila 2)
            for cell in worksheet[2]:
                if cell.value == col_name:
                    col_idx = cell.column
                    # Aplicar formato desde fila 3 (datos)
                    for row in range(3, worksheet.max_row + 1):
                        worksheet.cell(row=row, column=col_idx).number_format = EXCEL_CONFIG['currency_format']
                    break

    # e) Formato de fecha corta para columnas datetime del df
    columnas_fecha = df.select_dtypes(include=['datetime64[ns]', 'datetime64[ns, UTC]']).columns.tolist()
    for col_name in columnas_fecha:
        # Ubicar columna
        for cell in worksheet[2]:
            if cell.value == col_name:
                col_idx = cell.column
                for row in range(3, worksheet.max_row + 1):
                    worksheet.cell(row=row, column=col_idx).number_format = EXCEL_CONFIG['date_format']
                break

    # f) Relleno azul en encabezados específicos de la hoja "Mora"
    if es_hoja_mora:
        for col_name in MORA_BLUE_COLUMNS:
            if col_name in df.columns:
                # Buscar y aplicar relleno azul al encabezado (fila 2)
                for cell in worksheet[2]:
                    if cell.value == col_name:
                        cell.fill = PatternFill(start_color=COLORS['light_blue'], end_color=COLORS['light_blue'], fill_type="solid")
                        break

    # g) Inmovilización de paneles en A3
    worksheet.freeze_panes = EXCEL_CONFIG['freeze_panes']

def aplicar_formato_condicional(worksheet, columna_mora, num_filas):
    """Aplica formato condicional de colores a la columna de días de mora"""
    color_scale_rule = ColorScaleRule(
        start_type='min', start_color='7AB800', # Verde
        mid_type='percentile', mid_value=50, mid_color='FFEB84', # Amarillo
        end_type='max', end_color='FF6464' # Rojo
    )
    
    # Encuentra la letra de la columna 'Días de mora' (ahora en fila 2)
    mora_col_letter = [col[0].column_letter for col in worksheet.iter_cols(min_row=2, max_row=2) if col[0].value == columna_mora][0]
    # Aplicar formato desde fila 3 (datos) hasta el final
    worksheet.conditional_formatting.add(f'{mora_col_letter}3:{mora_col_letter}{num_filas + 2}', color_scale_rule)

def crear_tabla_excel(worksheet, df, sheet_name):
    """
    Convierte un rango de datos en una tabla formal de Excel con 10 columnas adicionales y títulos combinados.
    
    Estructura:
    - Fila 1: Títulos superiores combinados y coloreados
    - Fila 2: Encabezados de la tabla (datos + 10 en blanco)
    - Fila 3+: Datos de la tabla
    
    Args:
        worksheet: El objeto de la hoja de Excel
        df: El DataFrame correspondiente con los datos
        sheet_name: El nombre de la hoja para crear un nombre único de tabla
    """
    try:
        from openpyxl.styles import PatternFill, Alignment
        
        # Calcular el rango de la tabla dinámicamente
        num_filas_datos = len(df)  # Número de filas de datos
        num_filas_tabla = num_filas_datos + 1  # +1 para incluir el encabezado
        num_columnas_originales = len(df.columns)
        num_columnas_totales = num_columnas_originales + 9  # +9 columnas adicionales
        
        # Obtener las letras de las columnas
        col_inicio = get_column_letter(1)  # Columna A
        col_fin_original = get_column_letter(num_columnas_originales)  # Última columna de datos
        col_fin_total = get_column_letter(num_columnas_totales)  # Última columna incluyendo las 9 adicionales
        
        # --- PASO 1: Crear títulos combinados con colores en la FILA 1 ---
        # Título 1 (Verde): "Seguimiento Call Center" - primeras 4 columnas adicionales
        titulo1_inicio = get_column_letter(num_columnas_originales + 1)
        titulo1_fin = get_column_letter(num_columnas_originales + ADDITIONAL_COLUMNS['titles']['green']['columns'])
        rango_titulo1 = f"{titulo1_inicio}1:{titulo1_fin}1"
        worksheet.merge_cells(rango_titulo1)
        
        celda_titulo1 = worksheet[f"{titulo1_inicio}1"]
        celda_titulo1.value = ADDITIONAL_COLUMNS['titles']['green']['text']
        celda_titulo1.fill = PatternFill(start_color=COLORS['green'], end_color=COLORS['green'], fill_type="solid")
        celda_titulo1.alignment = Alignment(horizontal="center", vertical="center")
        
        # Título 2 (Azul): "Gestión de Cobranza en Campo" - siguientes 5 columnas adicionales
        titulo2_inicio = get_column_letter(num_columnas_originales + ADDITIONAL_COLUMNS['titles']['green']['columns'] + 1)
        titulo2_fin = get_column_letter(num_columnas_originales + ADDITIONAL_COLUMNS['count'])
        rango_titulo2 = f"{titulo2_inicio}1:{titulo2_fin}1"
        worksheet.merge_cells(rango_titulo2)
        
        celda_titulo2 = worksheet[f"{titulo2_inicio}1"]
        celda_titulo2.value = ADDITIONAL_COLUMNS['titles']['blue']['text']
        celda_titulo2.fill = PatternFill(start_color=COLORS['blue'], end_color=COLORS['blue'], fill_type="solid")
        celda_titulo2.alignment = Alignment(horizontal="center", vertical="center")
        
        # --- PASO 2: Agregar encabezados específicos en la FILA 2 ---
        for i, encabezado in enumerate(ADDITIONAL_COLUMNS['headers']):
            col_letra = get_column_letter(num_columnas_originales + 1 + i)
            worksheet.cell(row=2, column=num_columnas_originales + 1 + i, value=encabezado)
        
        # --- PASO 3: Crear la tabla con el rango correcto (empezando en FILA 2) ---
        # Crear el rango de la tabla incluyendo las 9 columnas adicionales, empezando en fila 2
        rango_tabla = f"{col_inicio}2:{col_fin_total}{num_filas_tabla + 1}"  # +1 porque startrow=1 mueve todo una fila
        
        # Crear el objeto Table con nombre único y válido
        nombre_tabla = generate_valid_table_name(sheet_name)
        logger.info(f"Creando tabla con nombre: '{nombre_tabla}' para hoja: '{sheet_name}'")
        tabla = Table(displayName=nombre_tabla, ref=rango_tabla)
        
        # Aplicar estilo a la tabla
        estilo = TableStyleInfo(
            name=EXCEL_CONFIG['table_style'],
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=False,
            showColumnStripes=False
        )
        tabla.tableStyleInfo = estilo
        
        # Añadir la tabla a la hoja
        worksheet.add_table(tabla)
        
        # Ajustar automáticamente el ancho de las columnas para que se vea todo el texto
        for i in range(1, num_columnas_totales + 1):
            column_letter = get_column_letter(i)
            max_length = 0
            # Revisar desde fila 1 (títulos) hasta el final de los datos
            for row in range(1, num_filas_tabla + 2):  # +2 porque los datos empiezan en fila 2
                cell_value = worksheet.cell(row=row, column=i).value
                if cell_value:
                    max_length = max(max_length, len(str(cell_value)))
            
            adjusted_width = min(max_length + 2, 50)  # Máximo 50 caracteres de ancho
            worksheet.column_dimensions[column_letter].width = adjusted_width
        
    except Exception as e:
        # Si hay algún error, no interrumpir el proceso principal
        logger.warning(f"No se pudo crear la tabla para la hoja {sheet_name}: {str(e)}")

def procesar_reporte_antiguedad(archivo_path):
    """Procesa el reporte de antigüedad con mejoras de robustez y mantenibilidad"""
    try:
        # Validar archivo
        validate_file_size(archivo_path)
        
        # --- PASO 1: Cargar y limpiar ---
        logger.info(f"Iniciando procesamiento del archivo: {archivo_path}")
        df = pd.read_excel(archivo_path, engine='openpyxl', dtype=DTYPE_CONFIG)
        df = clean_dataframe_columns(df)
        
        # Verificación de integridad de datos - ANTES de transformaciones
        medio_comunic_1_antes = df['Medio comunic. 1'].notna().sum() if 'Medio comunic. 1' in df.columns else 0
        medio_comunic_2_antes = df['Medio comunic. 2'].notna().sum() if 'Medio comunic. 2' in df.columns else 0
        logger.info(f"Verificación de integridad - ANTES: 'Medio comunic. 1' -> {medio_comunic_1_antes}, 'Medio comunic. 2' -> {medio_comunic_2_antes}")
        
        # Obtener nombres de columnas desde configuración
        columna_codigo = COLUMN_MAPPING['codigo']
        columna_mora = COLUMN_MAPPING['mora']
        columna_coordinacion = COLUMN_MAPPING['coordinacion']
        columna_geolocalizacion = COLUMN_MAPPING['geolocalizacion']
        
        # Validar que las columnas requeridas existan
        required_columns = [columna_codigo, columna_mora, columna_coordinacion]
        missing_columns = [col for col in required_columns if col not in df.columns]
        if missing_columns:
            raise ValueError(f"Columnas requeridas no encontradas: {missing_columns}")
        # --- PASO 1.1: Limpieza de datos ---
        # Estandarizar códigos a 6 dígitos
        code_columns = [columna_codigo, 'Código promotor', 'Código recuperador']
        df = standardize_codes(df, code_columns)
        
        # Limpiar números de teléfono
        df = clean_phone_numbers(df)
        
        # Añadir enlaces de geolocalización
        df = add_geolocation_links(df, columna_geolocalizacion)

        # --- PASO 1.2: Crear Informe Completo (antes de filtrar) ---
        logger.info("Creando informe completo con todos los registros")
        df_completo = df.sort_values(by=columna_mora, ascending=False).copy()
        df_completo = add_par_column(df_completo, columna_mora)
        
        # Eliminar columna "PAR 2" duplicada si existe en informe completo
        if 'PAR 2' in df_completo.columns:
            df_completo = df_completo.drop(columns=['PAR 2'], errors='ignore')
            logger.info("Columna 'PAR 2' duplicada eliminada del informe completo")
        
        df_completo_sin_links = df_completo.drop(columns=['link_texto', 'link_url'], errors='ignore')

        # --- PASO 2: Filtrar fraudes ---
        logger.info(f"Filtrando {len(LISTA_FRAUDE)} códigos de fraude")
        df_filtrado = df[~df[columna_codigo].isin(LISTA_FRAUDE)]

        # --- PASO 3: Ordenar y añadir columnas calculadas ---
        df_ordenado = df_filtrado.sort_values(by=columna_mora, ascending=False)
        df_ordenado = add_par_column(df_ordenado, columna_mora)
        
        # Eliminar columna "PAR 2" duplicada si existe
        if 'PAR 2' in df_ordenado.columns:
            df_ordenado = df_ordenado.drop(columns=['PAR 2'], errors='ignore')
            logger.info("Columna 'PAR 2' duplicada eliminada")
        
        # Verificación de integridad de datos - DESPUÉS de transformaciones
        medio_comunic_1_despues = df_ordenado['Medio comunic. 1'].notna().sum() if 'Medio comunic. 1' in df_ordenado.columns else 0
        medio_comunic_2_despues = df_ordenado['Medio comunic. 2'].notna().sum() if 'Medio comunic. 2' in df_ordenado.columns else 0
        
        # Verificar integridad
        if medio_comunic_1_antes == medio_comunic_1_despues:
            logger.info(f"Verificación 'Medio comunic. 1': Antes -> {medio_comunic_1_antes}, Después -> {medio_comunic_1_despues}. OK.")
        else:
            logger.warning(f"Verificación 'Medio comunic. 1': Antes -> {medio_comunic_1_antes}, Después -> {medio_comunic_1_despues}. PÉRDIDA DE DATOS!")
            
        if medio_comunic_2_antes == medio_comunic_2_despues:
            logger.info(f"Verificación 'Medio comunic. 2': Antes -> {medio_comunic_2_antes}, Después -> {medio_comunic_2_despues}. OK.")
        else:
            logger.warning(f"Verificación 'Medio comunic. 2': Antes -> {medio_comunic_2_antes}, Después -> {medio_comunic_2_despues}. PÉRDIDA DE DATOS!")

        # --- PASO 4: Crear DataFrame de Mora ---
        df_mora = df_ordenado[df_ordenado[columna_mora] >= 1].copy()
        logger.info(f"Registros en mora: {len(df_mora)}")

        # --- PASO 5: Distribuir ---
        coordinaciones_data = {}
        lista_coordinaciones = df_ordenado[columna_coordinacion].unique()
        for coord in lista_coordinaciones:
            if pd.notna(coord):
                coordinaciones_data[coord] = df_ordenado[df_ordenado[columna_coordinacion] == coord].copy()

        # --- PASO 6: Generar el archivo Excel final ---
        fecha_actual = datetime.now().strftime("%d%m%Y")
        nombre_archivo_salida = f'ReportedeAntigüedad_{fecha_actual}.xlsx'
        ruta_salida = os.path.join('uploads', nombre_archivo_salida)
        
        with pd.ExcelWriter(ruta_salida, engine='openpyxl') as writer:
            # --- Hoja 0: Informe completo ---
            hoja_informe = fecha_actual
            df_completo_sin_links.to_excel(writer, sheet_name=hoja_informe, index=False, startrow=1)
            ws_informe = writer.sheets[hoja_informe]
            # Aplicar formato condicional a la hoja de informe completo
            aplicar_formato_condicional(ws_informe, columna_mora, len(df_completo))
            
            # Añadir hipervínculos si existe geolocalización en informe completo
            if 'link_texto' in df_completo.columns and columna_geolocalizacion in df_completo.columns:
                geo_index = df_completo.columns.get_loc(columna_geolocalizacion)
                link_col = geo_index + 1  # Columna después de geolocalización
                
                # Escribir encabezado
                ws_informe.cell(row=2, column=link_col, value='Link de Geolocalización')
                
                # Escribir hipervínculos
                for i, (idx, row) in enumerate(df_completo.iterrows()):
                    row_num = i + 3  # +3 porque Excel empieza en 1, hay títulos en fila 1, encabezados en fila 2, datos empiezan en fila 3
                    texto = row['link_texto']
                    url = row['link_url']
                    escribir_hipervinculo_excel(ws_informe, row_num, link_col, texto, url)
            
            # Crear tabla y aplicar formato final
            crear_tabla_excel(ws_informe, df_completo_sin_links, hoja_informe)
            aplicar_formato_final(ws_informe, df_completo_sin_links, es_hoja_mora=False)
            # --- PASO 6.1: Crear hoja "Mora" ---
            # Escribir datos (empezando en fila 2)
            df_mora_sin_links = df_mora.drop(columns=['link_texto', 'link_url'], errors='ignore')
            df_mora_sin_links.to_excel(writer, sheet_name='Mora', index=False, startrow=1)
            
            # Aplicar formato condicional
            worksheet_mora = writer.sheets['Mora']
            aplicar_formato_condicional(worksheet_mora, columna_mora, len(df_mora))
            
            # Añadir hipervínculos si existe geolocalización
            if 'link_texto' in df_mora.columns and columna_geolocalizacion in df_mora.columns:
                geo_index = df_mora.columns.get_loc(columna_geolocalizacion)
                link_col = geo_index + 1  # Columna después de geolocalización
                
                # Escribir encabezado
                worksheet_mora.cell(row=2, column=link_col, value='Link de Geolocalización')
                
                # Escribir hipervínculos
                for i, (idx, row) in enumerate(df_mora.iterrows()):
                    row_num = i + 3  # +3 porque Excel empieza en 1, hay títulos en fila 1, encabezados en fila 2, datos empiezan en fila 3
                    texto = row['link_texto']
                    url = row['link_url']
                    escribir_hipervinculo_excel(worksheet_mora, row_num, link_col, texto, url)
            
            # Crear tabla formal de Excel para la hoja Mora y formato final
            crear_tabla_excel(worksheet_mora, df_mora_sin_links, 'Mora')
            aplicar_formato_final(worksheet_mora, df_mora_sin_links, es_hoja_mora=True)

            # --- PASO 6.2: Crear hojas por coordinación ---
            for coord_name, df_coord in coordinaciones_data.items():
                sheet_name = coord_name.replace(' ', '_')[:31]
                # Escribir datos (empezando en fila 2)
                df_coord_sin_links = df_coord.drop(columns=['link_texto', 'link_url'], errors='ignore')
                df_coord_sin_links.to_excel(writer, sheet_name=sheet_name, index=False, startrow=1)
                
                # Aplicar formato condicional
                worksheet_coord = writer.sheets[sheet_name]
                aplicar_formato_condicional(worksheet_coord, columna_mora, len(df_coord))
                
                # Añadir hipervínculos si existe geolocalización
                if 'link_texto' in df_coord.columns and columna_geolocalizacion in df_coord.columns:
                    geo_index = df_coord.columns.get_loc(columna_geolocalizacion)
                    link_col = geo_index + 1  # Columna después de geolocalización
                    
                    # Escribir encabezado
                    worksheet_coord.cell(row=2, column=link_col, value='Link de Geolocalización')
                    
                    # Escribir hipervínculos
                    for i, (idx, row) in enumerate(df_coord.iterrows()):
                        row_num = i + 3  # +3 porque Excel empieza en 1, hay títulos en fila 1, encabezados en fila 2, datos empiezan en fila 3
                        texto = row['link_texto']
                        url = row['link_url']
                        escribir_hipervinculo_excel(worksheet_coord, row_num, link_col, texto, url)

                # Crear tabla formal de Excel para la hoja de coordinación y formato final
                crear_tabla_excel(worksheet_coord, df_coord_sin_links, sheet_name)
                aplicar_formato_final(worksheet_coord, df_coord_sin_links, es_hoja_mora=False)

        logger.info(f"Procesamiento completado exitosamente. Archivo generado: {ruta_salida}")
        return ruta_salida, len(coordinaciones_data)
        
    except FileNotFoundError as e:
        logger.error(f"Archivo no encontrado: {str(e)}")
        raise Exception(f"El archivo especificado no existe: {str(e)}")
    except pd.errors.EmptyDataError as e:
        logger.error(f"Archivo Excel vacío: {str(e)}")
        raise Exception(f"El archivo Excel está vacío o no contiene datos válidos: {str(e)}")
    except pd.errors.ExcelFileError as e:
        logger.error(f"Error al leer archivo Excel: {str(e)}")
        raise Exception(f"Error al leer el archivo Excel. Verifique que sea un archivo válido: {str(e)}")
    except ValueError as e:
        logger.error(f"Error de validación: {str(e)}")
        raise Exception(f"Error de validación de datos: {str(e)}")
    except Exception as e:
        logger.error(f"Error inesperado procesando archivo: {str(e)}")
        raise Exception(f"Error inesperado procesando archivo: {str(e)}")

@reportes_bp.route('/antiguedad')
def antiguedad_form():
    """Página para subir archivo de reporte de antigüedad"""
    return render_template('antiguedad.html')

@reportes_bp.route('/download/<filename>')
def download_file(filename):
    """Servir archivos de descarga"""
    file_path = os.path.join('static', 'downloads', filename)
    if os.path.exists(file_path):
        return send_file(file_path, as_attachment=True)
    else:
        return "Archivo no encontrado", 404

@reportes_bp.route('/antiguedad/procesar', methods=['POST'])
def procesar_antiguedad():
    """Procesa el archivo subido y devuelve el reporte"""
    if 'archivo' not in request.files:
        flash('No se seleccionó ningún archivo', 'error')
        return redirect(url_for('reportes.antiguedad_form'))
    
    archivo = request.files['archivo']
    if archivo.filename == '':
        flash('No se seleccionó ningún archivo', 'error')
        return redirect(url_for('reportes.antiguedad_form'))
    
    if not allowed_file(archivo.filename):
        flash('El archivo debe ser de tipo Excel (.xlsx o .xls)', 'error')
        return redirect(url_for('reportes.antiguedad_form'))
    
    try:
        # Validar tamaño del archivo
        archivo.seek(0, 2)  # Ir al final del archivo
        file_size = archivo.tell()
        archivo.seek(0)  # Volver al inicio
        
        if file_size > MAX_FILE_SIZE:
            flash(f'El archivo es demasiado grande. Tamaño máximo permitido: {MAX_FILE_SIZE // (1024*1024)}MB', 'error')
            return redirect(url_for('reportes.antiguedad_form'))
        
        # Guardar archivo temporalmente
        filename = secure_filename(archivo.filename)
        archivo_path = os.path.join(UPLOAD_FOLDER, filename)
        archivo.save(archivo_path)
        
        logger.info(f"Archivo subido exitosamente: {filename} ({file_size} bytes)")
        
        # Procesar archivo
        ruta_salida, num_coordinaciones = procesar_reporte_antiguedad(archivo_path)
        
        # Limpiar archivo temporal
        try:
            os.remove(archivo_path)
        except (OSError, FileNotFoundError):
            pass  # Ignorar errores de eliminación
        
        # Guardar archivo en carpeta de descargas
        import shutil
        download_path = os.path.join('static', 'downloads')
        os.makedirs(download_path, exist_ok=True)
        shutil.copy2(ruta_salida, os.path.join(download_path, os.path.basename(ruta_salida)))
        
        # Usar flash para mostrar mensaje de éxito con enlace de descarga
        flash(f'Reporte procesado exitosamente. <a href="{url_for("reportes.download_file", filename=os.path.basename(ruta_salida))}" class="btn btn-sm btn-primary ms-2">Descargar</a>', 'success')
        
        # Redirigir de vuelta al formulario
        return redirect(url_for('reportes.antiguedad_form'))
        
    except Exception as e:
        logger.error(f"Error en procesamiento de archivo: {str(e)}")
        flash(f'Error procesando archivo: {str(e)}', 'error')
        return redirect(url_for('reportes.antiguedad_form'))
