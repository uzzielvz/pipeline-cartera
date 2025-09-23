from flask import Blueprint, render_template, request, send_file, flash, redirect, url_for, Response
import pandas as pd
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.styles import Font
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter
from openpyxl.styles import numbers
import os
import re
from datetime import datetime
from werkzeug.utils import secure_filename
import urllib.parse

reportes_bp = Blueprint('reportes', __name__)

# Lista de fraude (exactamente igual al pipeline original)
LISTA_FRAUDE = [
    "001041", "001005", "001023", "001018", "001014", "001024", "001025", "001042",
    "001019", "001026", "001048", "001049", "001050", "001051", "001028", "001002",
    "001008", "001034", "001010", "001045", "001044", "001029", "001007", "001032",
    "001022", "001000", "001040"
]

def allowed_file(filename):
    """Verifica que el archivo sea Excel"""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in {'xlsx', 'xls'}

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

def aplicar_formato_final(worksheet, df):
    """Autoajuste de columnas, formato de moneda y fecha corta."""
    # a) Autoajuste de columnas
    for i in range(1, worksheet.max_column + 1):
        column_letter = get_column_letter(i)
        max_length = 0
        for j in range(1, worksheet.max_row + 1):
            value = worksheet.cell(row=j, column=i).value
            if value is not None:
                max_length = max(max_length, len(str(value)))
        worksheet.column_dimensions[column_letter].width = min(max_length + 2, 50)

    # b) Formato de moneda en columnas conocidas del df
    columnas_moneda = [
        col for col in df.columns
        if any(key in col.lower() for key in ['monto', 'saldo', 'importe', 'cantidad', 'pago'])
    ]
    for col_name in columnas_moneda:
        if col_name in df.columns:
            # Buscar índice de la columna por encabezado (fila 2)
            for cell in worksheet[2]:
                if cell.value == col_name:
                    col_idx = cell.column
                    # Aplicar formato desde fila 3 (datos)
                    for row in range(3, worksheet.max_row + 1):
                        worksheet.cell(row=row, column=col_idx).number_format = "$#,##0.00"
                    break

    # c) Formato de fecha corta para columnas datetime del df
    columnas_fecha = df.select_dtypes(include=['datetime64[ns]', 'datetime64[ns, UTC]']).columns.tolist()
    for col_name in columnas_fecha:
        # Ubicar columna
        for cell in worksheet[2]:
            if cell.value == col_name:
                col_idx = cell.column
                for row in range(3, worksheet.max_row + 1):
                    worksheet.cell(row=row, column=col_idx).number_format = "DD/MM/YYYY"
                break

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
        num_columnas_totales = num_columnas_originales + 10  # +10 columnas adicionales
        
        # Obtener las letras de las columnas
        col_inicio = get_column_letter(1)  # Columna A
        col_fin_original = get_column_letter(num_columnas_originales)  # Última columna de datos
        col_fin_total = get_column_letter(num_columnas_totales)  # Última columna incluyendo las 10 adicionales
        
        # --- PASO 1: Crear títulos combinados con colores en la FILA 1 ---
        # Distribuir uniformemente los títulos entre las 10 columnas adicionales
        # Título 1 (Verde): "Información proporcionada por el gerente de" - primeras 5 columnas adicionales
        titulo1_inicio = get_column_letter(num_columnas_originales + 1)
        titulo1_fin = get_column_letter(num_columnas_originales + 5)
        rango_titulo1 = f"{titulo1_inicio}1:{titulo1_fin}1"
        worksheet.merge_cells(rango_titulo1)
        
        celda_titulo1 = worksheet[f"{titulo1_inicio}1"]
        celda_titulo1.value = "Información proporcionada por el gerente de"
        celda_titulo1.fill = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")  # Verde claro
        celda_titulo1.alignment = Alignment(horizontal="center", vertical="center")
        
        # Título 2 (Azul): "Gestión del auxiliar de operaciones" - siguientes 5 columnas adicionales
        titulo2_inicio = get_column_letter(num_columnas_originales + 6)
        titulo2_fin = get_column_letter(num_columnas_originales + 10)
        rango_titulo2 = f"{titulo2_inicio}1:{titulo2_fin}1"
        worksheet.merge_cells(rango_titulo2)
        
        celda_titulo2 = worksheet[f"{titulo2_inicio}1"]
        celda_titulo2.value = "Gestión del auxiliar de operaciones"
        celda_titulo2.fill = PatternFill(start_color="87CEEB", end_color="87CEEB", fill_type="solid")  # Azul claro
        celda_titulo2.alignment = Alignment(horizontal="center", vertical="center")
        
        # --- PASO 2: Agregar 10 encabezados completamente en blanco en la FILA 2 ---
        for i in range(10):
            col_letra = get_column_letter(num_columnas_originales + 1 + i)
            worksheet.cell(row=2, column=num_columnas_originales + 1 + i, value="")
        
        # --- PASO 3: Crear la tabla con el rango correcto (empezando en FILA 2) ---
        # Crear el rango de la tabla incluyendo las 10 columnas adicionales, empezando en fila 2
        rango_tabla = f"{col_inicio}2:{col_fin_total}{num_filas_tabla + 1}"  # +1 porque startrow=1 mueve todo una fila
        
        # Crear el objeto Table con nombre único
        nombre_tabla = f"Tabla_{sheet_name.replace(' ', '_')}"
        tabla = Table(displayName=nombre_tabla, ref=rango_tabla)
        
        # Aplicar estilo a la tabla
        estilo = TableStyleInfo(
            name="TableStyleLight1",
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
        print(f"Advertencia: No se pudo crear la tabla para la hoja {sheet_name}: {str(e)}")

def procesar_reporte_antiguedad(archivo_path):
    """Lógica exacta del pipeline.py original con funcionalidad de macro CopiarMora"""
    try:
        # --- PASO 1: Cargar y limpiar ---
        df = pd.read_excel(archivo_path, engine='openpyxl')
        df.columns = df.columns.str.replace('\n', ' ').str.strip()
        columna_codigo = 'Código acreditado'
        columna_mora = 'Días de mora'
        columna_coordinacion = 'Coordinación'
        columna_geolocalizacion = 'Geolocalización domicilio'
        # Estandarizar códigos a 6 dígitos (texto con ceros a la izquierda)
        df[columna_codigo] = pd.to_numeric(df[columna_codigo], errors='coerce').fillna(0).astype(int).astype(str).str.zfill(6)
        for col_codigo in ['Código promotor', 'Código recuperador']:
            if col_codigo in df.columns:
                df[col_codigo] = pd.to_numeric(df[col_codigo], errors='coerce').fillna(0).astype(int).astype(str).str.zfill(6)

        # Manejo de datos faltantes: teléfonos como strings vacíos
        for col in df.columns:
            if 'Teléfono' in col:
                df[col] = df[col].fillna('').astype(str)
        
        # --- PASO 1.1: Procesar geolocalización ---
        if columna_geolocalizacion in df.columns:
            # Aplicar función de geolocalización
            geolocalizacion_data = df[columna_geolocalizacion].apply(generar_link_google_maps)
            df['link_texto'] = [item[0] for item in geolocalizacion_data]
            df['link_url'] = [item[1] for item in geolocalizacion_data]

        # --- PASO 1.2: Crear Informe Completo (antes de filtrar) ---
        # Ordenar por días de mora (incluye clientes con 0 días)
        df_completo = df.sort_values(by=columna_mora, ascending=False).copy()
        df_completo_sin_links = df_completo.drop(columns=['link_texto', 'link_url'], errors='ignore')

        # --- PASO 2: Filtrar ---
        df_filtrado = df[~df[columna_codigo].isin(LISTA_FRAUDE)]

        # --- PASO 3: Ordenar ---
        df_ordenado = df_filtrado.sort_values(by=columna_mora, ascending=False)

        # --- PASO 4: Crear DataFrame de Mora ---
        df_mora = df_ordenado[df_ordenado[columna_mora] >= 1].copy()
        
        # --- PASO 4.1: Añadir columna PAR ---
        df_mora['PAR'] = df_mora[columna_mora].apply(asignar_rango_mora)
        
        # --- PASO 4.2: Reordenar columnas para que 'PAR' esté al lado de 'Días de mora' ---
        columnas = df_mora.columns.tolist()
        mora_index = columnas.index(columna_mora)
        
        # Crear nuevo orden: insertar 'PAR' después de 'Días de mora'
        nuevas_columnas = (columnas[:mora_index+1] + 
                          ['PAR'] + 
                          columnas[mora_index+1:])
        
        df_mora = df_mora[nuevas_columnas]

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
            # Crear tabla y aplicar formato final
            crear_tabla_excel(ws_informe, df_completo_sin_links, hoja_informe)
            aplicar_formato_final(ws_informe, df_completo_sin_links)
            # --- PASO 6.1: Crear hoja "Mora" (replicando macro CopiarMora) ---
            # Escribir datos sin las columnas temporales de links (empezando en fila 2)
            df_mora_sin_links = df_mora.drop(columns=['link_texto', 'link_url'], errors='ignore')
            df_mora_sin_links.to_excel(writer, sheet_name='Mora', index=False, startrow=1)
            
            # Aplicar formato condicional
            worksheet_mora = writer.sheets['Mora']
            aplicar_formato_condicional(worksheet_mora, columna_mora, len(df_mora))
            
            # Añadir columna de hipervínculos si existe geolocalización
            if 'link_texto' in df_mora.columns:
                # Encontrar la posición de la columna geolocalización original
                if columna_geolocalizacion in df_mora.columns:
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
            aplicar_formato_final(worksheet_mora, df_mora_sin_links)

            # --- PASO 6.2: Crear hojas por coordinación ---
            for coord_name, df_coord in coordinaciones_data.items():
                sheet_name = coord_name.replace(' ', '_')[:31]
                # Escribir datos sin las columnas temporales de links (empezando en fila 2)
                df_coord_sin_links = df_coord.drop(columns=['link_texto', 'link_url'], errors='ignore')
                df_coord_sin_links.to_excel(writer, sheet_name=sheet_name, index=False, startrow=1)
                
                # Aplicar formato condicional
                worksheet_coord = writer.sheets[sheet_name]
                aplicar_formato_condicional(worksheet_coord, columna_mora, len(df_coord))
                
                # Añadir columna de hipervínculos si existe geolocalización
                if 'link_texto' in df_coord.columns:
                    if columna_geolocalizacion in df_coord.columns:
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
                aplicar_formato_final(worksheet_coord, df_coord_sin_links)

        return ruta_salida, len(coordinaciones_data)
        
    except Exception as e:
        raise Exception(f"Error procesando archivo: {str(e)}")

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
        # Guardar archivo temporalmente
        filename = secure_filename(archivo.filename)
        archivo_path = os.path.join('uploads', filename)
        archivo.save(archivo_path)
        
        # Procesar archivo
        ruta_salida, num_coordinaciones = procesar_reporte_antiguedad(archivo_path)
        
        # Limpiar archivo temporal
        os.remove(archivo_path)
        
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
        flash(f'Error procesando archivo: {str(e)}', 'error')
        return redirect(url_for('reportes.antiguedad_form'))
