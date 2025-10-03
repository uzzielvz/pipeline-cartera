"""
Configuración centralizada para el sistema de reportes de antigüedad
"""

# Configuración de archivos
ALLOWED_EXTENSIONS = {'xlsx', 'xls'}
UPLOAD_FOLDER = 'uploads'
MAX_FILE_SIZE = 16 * 1024 * 1024  # 16MB

# Configuración de columnas
COLUMN_MAPPING = {
    'codigo': 'Código acreditado',
    'mora': 'Días de mora',
    'coordinacion': 'Coordinación',
    'geolocalizacion': 'Geolocalización domicilio',
    'saldo_vencido': 'Saldo vencido',
    'ciclo': 'Ciclo',
    'nombre_acreditado': 'Nombre acreditado',
    'intereses_vencidos': 'Intereses vencidos',
    'comision_vencida': 'Comisión vencida',
    'recargos': 'Recargos',
    'saldo_capital': 'Saldo capital'
}

# Configuración de tipos de datos para prevenir pérdida de datos
DTYPE_CONFIG = {
    'Medio comunic. 1': str,
    'Medio comunic. 2': str,
    'Medio comunic. 3': str,
    'Teléfono conyuge': str,
    'Teléfono Referencia1': str,
    'Teléfono Referencia2': str,
    'Teléfono Referencia3': str
}

# Lista de códigos de fraude
LISTA_FRAUDE = [
    "001041", "001005", "001023", "001018", "001014", "001024", "001025", "001042",
    "001019", "001026", "001048", "001049", "001050", "001051", "001028", "001002",
    "001008", "001034", "001010", "001045", "001044", "001029", "001007", "001032",
    "001022", "001000", "001040"
]

# Configuración de formato Excel
EXCEL_CONFIG = {
    'table_style': 'TableStyleLight1',
    'header_height': 35,
    'freeze_panes': 'A3',
    'max_column_width': 50,
    'currency_format': '$#,##0.00',
    'date_format': 'DD/MM/YYYY'
}

# Configuración de colores
COLORS = {
    'light_blue': 'DDEBF7',
    'green': 'C6EFCE',
    'blue': 'B3D9FF',
    'light_green': 'E6FFE6'
}

# Configuración de columnas adicionales
ADDITIONAL_COLUMNS = {
    'count': 9,
    'headers': [
        "Estatus de llamada (pago del día ó mora)",
        "Fecha del acuerdo de pago",
        "Horario del acuerdo de pago",
        "Monto del acuerdo",
        "Día de visita de cobranza en campo",
        "Fecha del acuerdo de pago",
        "Horario del acuerdo de pago",
        "Monto del acuerdo",
        "Monto del acuerdo"
    ],
    'titles': {
        'green': {
            'text': 'Seguimiento Call Center',
            'columns': 4
        },
        'blue': {
            'text': 'Gestión de Cobranza en Campo',
            'columns': 5
        }
    }
}

# Configuración de columnas con relleno azul en hoja Mora
MORA_BLUE_COLUMNS = [
    'Saldo capital',
    'Saldo capital vencido',
    'Saldo interés vencido',
    'Saldo comisión vencida',
    'Saldo recargos'
]

# Configuración de columnas de moneda
CURRENCY_COLUMNS_KEYWORDS = ['monto', 'saldo', 'importe', 'cantidad', 'pago']

# Configuración de columnas de fecha
DATE_COLUMNS_KEYWORDS = ['fecha', 'date']

# Configuración de autenticación
SECRET_KEY = 'tu-clave-secreta-super-segura-aqui-cambiar-en-produccion'
SQLALCHEMY_DATABASE_URI = 'sqlite:///crediflexi.db'
SQLALCHEMY_TRACK_MODIFICATIONS = False

# Configuración de roles
USER_ROLES = {
    'ADMIN': 'Administrador/Generador',
    'CONSULTOR': 'Consultor'
}

# Configuración de permisos
PERMISSIONS = {
    'ADMIN': ['generate_reports', 'view_reports', 'manage_users'],
    'CONSULTOR': ['view_reports']
}
