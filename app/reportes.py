from flask import Blueprint, render_template, request, send_file, flash, redirect, url_for, Response
import pandas as pd
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.styles import Font
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
    - Coordenadas GPS: (coordenadas_texto, url_búsqueda)
    - Dirección de texto: (dirección_texto, url_búsqueda)
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
                return (geolocalizacion, url)
        except:
            pass
    
    # Caso 4: Dirección de texto - crear búsqueda
    direccion_encoded = urllib.parse.quote_plus(geolocalizacion)
    url = f"https://www.google.com/maps/search/?api=1&query={direccion_encoded}"
    return (geolocalizacion, url)

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

def aplicar_formato_condicional(worksheet, columna_mora, num_filas):
    """Aplica formato condicional de colores a la columna de días de mora"""
    color_scale_rule = ColorScaleRule(
        start_type='min', start_color='7AB800', # Verde
        mid_type='percentile', mid_value=50, mid_color='FFEB84', # Amarillo
        end_type='max', end_color='FF6464' # Rojo
    )
    
    # Encuentra la letra de la columna 'Días de mora'
    mora_col_letter = [col[0].column_letter for col in worksheet.iter_cols(min_row=1, max_row=1) if col[0].value == columna_mora][0]
    worksheet.conditional_formatting.add(f'{mora_col_letter}2:{mora_col_letter}{num_filas + 1}', color_scale_rule)

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
        df[columna_codigo] = pd.to_numeric(df[columna_codigo], errors='coerce').fillna(0).astype(int).astype(str).str.zfill(6)
        
        # --- PASO 1.1: Procesar geolocalización ---
        if columna_geolocalizacion in df.columns:
            # Aplicar función de geolocalización
            geolocalizacion_data = df[columna_geolocalizacion].apply(generar_link_google_maps)
            df['link_texto'] = [item[0] for item in geolocalizacion_data]
            df['link_url'] = [item[1] for item in geolocalizacion_data]

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
            # --- PASO 6.1: Crear hoja "Mora" (replicando macro CopiarMora) ---
            # Escribir datos sin las columnas temporales de links
            df_mora_sin_links = df_mora.drop(columns=['link_texto', 'link_url'], errors='ignore')
            df_mora_sin_links.to_excel(writer, sheet_name='Mora', index=False)
            
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
                    worksheet_mora.cell(row=1, column=link_col, value='Link de Geolocalización')
                    
                    # Escribir hipervínculos
                    for idx, row in df_mora.iterrows():
                        row_num = idx + 2  # +2 porque Excel empieza en 1 y hay encabezado
                        texto = row['link_texto']
                        url = row['link_url']
                        escribir_hipervinculo_excel(worksheet_mora, row_num, link_col, texto, url)

            # --- PASO 6.2: Crear hojas por coordinación ---
            for coord_name, df_coord in coordinaciones_data.items():
                sheet_name = coord_name.replace(' ', '_')[:31]
                # Escribir datos sin las columnas temporales de links
                df_coord_sin_links = df_coord.drop(columns=['link_texto', 'link_url'], errors='ignore')
                df_coord_sin_links.to_excel(writer, sheet_name=sheet_name, index=False)
                
                # Aplicar formato condicional
                worksheet_coord = writer.sheets[sheet_name]
                aplicar_formato_condicional(worksheet_coord, columna_mora, len(df_coord))
                
                # Añadir columna de hipervínculos si existe geolocalización
                if 'link_texto' in df_coord.columns:
                    if columna_geolocalizacion in df_coord.columns:
                        geo_index = df_coord.columns.get_loc(columna_geolocalizacion)
                        link_col = geo_index + 1  # Columna después de geolocalización
                        
                        # Escribir encabezado
                        worksheet_coord.cell(row=1, column=link_col, value='Link de Geolocalización')
                        
                        # Escribir hipervínculos
                        for idx, row in df_coord.iterrows():
                            row_num = idx + 2  # +2 porque Excel empieza en 1 y hay encabezado
                            texto = row['link_texto']
                            url = row['link_url']
                            escribir_hipervinculo_excel(worksheet_coord, row_num, link_col, texto, url)

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
