import pandas as pd
from openpyxl.formatting.rule import ColorScaleRule

# --- CONFIGURACIÓN INICIAL ---
LISTA_FRAUDE = [
    "001041", "001005", "001023", "001018", "001014", "001024", "001025", "001042",
    "001019", "001026", "001048", "001049", "001050", "001051", "001028", "001002",
    "001008", "001034", "001010", "001045", "001044", "001029", "001007", "001032",
    "001022", "001000", "001040"
]
nombre_archivo_entrada = 'Reporte en bruto.xlsx'
nombre_archivo_salida = 'Reporte_Final_Lunes.xlsx'

try:
    # --- PASO 1: Cargar, Limpiar y Convertir Tipos ---
    print("Paso 1: Cargando y limpiando datos...")
    df = pd.read_excel(nombre_archivo_entrada, engine='openpyxl')
    df.columns = df.columns.str.replace('\n', ' ').str.strip()
    
    columna_codigo = 'Código acreditado'
    columna_mora = 'Días de mora'
    columna_coordinacion = 'Coordinación'
    
    # Estandarizar columna de código
    df[columna_codigo] = pd.to_numeric(df[columna_codigo], errors='coerce').fillna(0).astype(int).astype(str).str.zfill(6)

    # Convertir columnas de moneda a números
    columnas_moneda = ['Cantidad entregada', 'Cantidad Prestada', 'Saldo capital', 'Saldo vencido', 'Saldo total']
    for col in columnas_moneda:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col].astype(str).str.replace(',', ''), errors='coerce').fillna(0)

    # Convertir columnas de fecha
    columnas_fecha = ['Inicio ciclo', 'Fin ciclo', 'Último pago']
    for col in columnas_fecha:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], format='%d/%m/%Y', errors='coerce')
    print("-> Carga y conversión de tipos completada.")

    # --- PASO 2: Filtrar ---
    print("Paso 2: Filtrando cartera de fraude...")
    df_filtrado = df[~df[columna_codigo].isin(LISTA_FRAUDE)]
    print(f"-> Se eliminaron {len(df) - len(df_filtrado)} filas.")

    # --- PASO 3: Ordenar ---
    print("Paso 3: Ordenando datos por riesgo...")
    df_ordenado = df_filtrado.sort_values(by=columna_mora, ascending=False)
    print("-> Datos ordenados.")

    # --- PASO 4: Distribuir ---
    print("Paso 4: Distribuyendo por coordinación...")
    coordinaciones_data = {}
    lista_coordinaciones = df_ordenado[columna_coordinacion].unique()
    for coord in lista_coordinaciones:
        if pd.notna(coord):
            coordinaciones_data[coord] = df_ordenado[df_ordenado[columna_coordinacion] == coord].copy()
    print("-> Datos distribuidos.")

    # --- PASO 5: Generar el archivo Excel final ---
    print(f"\n--- Creando archivo final: {nombre_archivo_salida} ---")
    with pd.ExcelWriter(nombre_archivo_salida, engine='openpyxl') as writer:
        for coord_name, df_coord in coordinaciones_data.items():
            sheet_name = coord_name.replace(' ', '_')[:31]
            df_coord.to_excel(writer, sheet_name=sheet_name, index=False)
            
            worksheet = writer.sheets[sheet_name]
            color_scale_rule = ColorScaleRule(start_type='min', start_color='7AB800', mid_type='percentile', mid_value=50, mid_color='FFEB84', end_type='max', end_color='FF6464')
            
            mora_col_letter = [col[0].column_letter for col in worksheet.iter_cols(min_row=1, max_row=1) if col[0].value == columna_mora][0]
            worksheet.conditional_formatting.add(f'{mora_col_letter}2:{mora_col_letter}{len(df_coord) + 1}', color_scale_rule)

    print(f"\n¡MISIÓN CUMPLIDA! Se ha generado el reporte final en tu carpeta.")

except Exception as e:
    print(f"Ocurrió un error: {e}")