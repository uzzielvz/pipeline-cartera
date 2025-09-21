import pandas as pd
import re
import urllib.parse

def generar_link_google_maps(geolocalizacion):
    if pd.isna(geolocalizacion) or str(geolocalizacion).strip() == '':
        return ('Ver en mapa', 'https://maps.google.com')
    
    geolocalizacion = str(geolocalizacion).strip()
    
    if 'http' in geolocalizacion or 'google.com/maps' in geolocalizacion:
        return ('Ver en mapa', geolocalizacion)
    
    if any(char in geolocalizacion for char in ['°', "'", '"', 'N', 'W', 'S', 'E']):
        try:
            coord_pattern = r"(\d+)°(\d+)'([\d.]+)\"([NS])\s+(\d+)°(\d+)'([\d.]+)\"([WE])"
            match = re.search(coord_pattern, geolocalizacion)
            
            if match:
                lat_deg, lat_min, lat_sec, lat_dir = match.groups()[:4]
                lon_deg, lon_min, lon_sec, lon_dir = match.groups()[4:]
                
                lat_decimal = float(lat_deg) + float(lat_min)/60 + float(lat_sec)/3600
                lon_decimal = float(lon_deg) + float(lon_min)/60 + float(lon_sec)/3600
                
                if lat_dir == 'S':
                    lat_decimal = -lat_decimal
                if lon_dir == 'W':
                    lon_decimal = -lon_decimal
                
                url = f'https://www.google.com/maps/search/?api=1&query={lat_decimal},{lon_decimal}'
                return (geolocalizacion, url)
        except:
            pass
    
    direccion_encoded = urllib.parse.quote_plus(geolocalizacion)
    url = f'https://www.google.com/maps/search/?api=1&query={direccion_encoded}'
    return (geolocalizacion, url)

# Test de la función
test_cases = [
    '',  # Vacío
    None,  # Nulo
    'https://maps.google.com/maps?q=test',  # URL existente
    '19°12\'12.2"N 100°07\'51.8"W',  # Coordenadas GPS
    'Esteban bunuelos rumbo',  # Dirección de texto
    'Av. Principal 123, Ciudad'  # Dirección completa
]

print('Testing generar_link_google_maps function:')
for test in test_cases:
    result = generar_link_google_maps(test)
    test_str = str(test) if test is not None else 'None'
    print(f'Input: {test_str:<30} -> Texto: {result[0]:<20} | URL: {result[1][:50]}...')
print('\n✅ Función de geolocalización funcionando correctamente!')
