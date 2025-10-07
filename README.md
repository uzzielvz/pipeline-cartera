# 🏦 CREDIFLEXI - Sistema de Automatización de Reportes

## 📋 Descripción General

CREDIFLEXI es una aplicación web profesional diseñada para automatizar el procesamiento de reportes de antigüedad de cartera crediticia. El sistema proporciona procesamiento automatizado de datos, filtrado de fraudes, integración de geolocalización, generación de reportes Excel profesionales con formato avanzado y estilos condicionales.

---

## ✨ Características Principales

### 🎯 Funcionalidad Core

#### 📊 Reporte de Antigüedad Individual
- Procesamiento automatizado de archivos Excel con datos de cartera crediticia
- Filtrado automático de códigos fraudulentos usando listas predefinidas
- Generación de enlaces de geolocalización en Google Maps
- Cálculo automático de PAR (Portfolio at Risk) basado en días de mora
- **NUEVO**: Opción para generar reporte separado de colaboradores (códigos 001053 y 001145)

#### 👥 Reporte de Antigüedad Grupal
- Integración de 5 archivos Excel para generar reporte consolidado
- Detección automática del tipo de cada archivo por contenido
- Procesamiento unificado de datos de múltiples fuentes

#### 🔐 Sistema de Usuarios
- **Administrador/Generador**: Puede generar reportes y acceder al dashboard
- **Consultor**: Solo puede consultar reportes existentes en el dashboard
- Autenticación segura con Flask-Login
- Gestión de permisos por rol

#### 📈 Dashboard Interactivo
- Visualización de estadísticas de reportes generados
- Filtros por tipo de reporte (Individual, Grupal, Colaboradores)
- Historial completo de reportes con detalles (fecha, usuario, tamaño)
- Descarga y eliminación de reportes (solo admin)

### 📑 Generación de Reportes

Cada reporte individual genera las siguientes hojas:

1. **Informe Completo** (fecha del día)
   - Dataset completo con todos los registros procesados
   - Código acreditado como primera columna
   - Enlaces de geolocalización clickeables

2. **Mora**
   - Registros con 1+ días de mora
   - Columnas adicionales para seguimiento:
     - **Seguimiento Call Center** (4 columnas verdes)
     - **Gestión de Cobranza en Campo** (5 columnas azules)

3. **Cuentas con saldo vencido**
   - Registros con saldo vencido ≥ 1 y días de mora ≤ 0

4. **Liquidación anticipada**
   - Calculadora interactiva con fórmulas VLOOKUP
   - Ingreso manual de código acreditado
   - Cálculo automático de montos para liquidación

5. **Hojas por Coordinación**
   - Una hoja por cada coordinación/región
   - Datos filtrados y organizados por área

### 🛠️ Capacidades de Procesamiento

- **Limpieza de Datos**: Estandariza números telefónicos y formatos
- **Mapeo de Columnas**: Detección inteligente de columnas
- **Prevención de Duplicados**: Manejo automático de columnas duplicadas (PAR)
- **Integridad de Datos**: Validación exhaustiva

---

## 🏗️ Arquitectura Técnica

### 💻 Stack Tecnológico

- **Backend**: Python 3.13 con Flask
- **Procesamiento de Datos**: Pandas
- **Generación Excel**: OpenPyXL
- **Base de Datos**: SQLite con SQLAlchemy
- **Autenticación**: Flask-Login
- **Formularios**: Flask-WTF
- **Frontend**: HTML5, CSS3, JavaScript (Vanilla)

### 📁 Estructura del Proyecto

```
automatizador-crediflexi/
├── app/
│   ├── __init__.py
│   ├── reportes.py          # Lógica de procesamiento de reportes
│   ├── auth.py              # Sistema de autenticación
│   ├── consultor.py         # Dashboard y consultas
│   └── models.py            # Modelos de base de datos
├── static/
│   ├── css/
│   │   └── style.css        # Estilos de la aplicación
│   ├── downloads/
│   │   └── reports/         # Reportes generados (permanentes)
│   └── js/                  # JavaScript
├── templates/
│   ├── base.html            # Template base
│   ├── index.html           # Página principal
│   ├── auth/
│   │   └── login.html       # Login de usuarios
│   ├── consultor/
│   │   ├── dashboard.html   # Dashboard de reportes
│   │   ├── reports.html     # Lista completa
│   │   └── report_detail.html
│   └── errors/
│       └── unauthorized.html
├── uploads/                 # Archivos temporales
├── instance/
│   └── crediflexi.db        # Base de datos SQLite
├── app.py                   # Punto de entrada de Flask
├── config.py               # Configuración general
├── .gitignore              # Archivos ignorados por Git
├── requirements.txt        # Dependencias Python
└── README.md              # Este archivo
```

---

## 🚀 Instalación Local - Guía Paso a Paso

### 📋 Prerrequisitos

Antes de comenzar, asegúrate de tener instalado:

1. **Python 3.13 o superior**
   - Descargar de: https://www.python.org/downloads/
   - ✅ Durante la instalación, marca "Add Python to PATH"

2. **Git** (para clonar el repositorio)
   - Descargar de: https://git-scm.com/downloads

3. **Editor de código** (recomendado)
   - Visual Studio Code: https://code.visualstudio.com/
   - O cualquier editor de tu preferencia

---

### 📥 Paso 1: Clonar el Repositorio

Abre una terminal (PowerShell en Windows, Terminal en Mac/Linux) y ejecuta:

```bash
# Clonar el repositorio
git clone <URL_DEL_REPOSITORIO>

# Entrar al directorio del proyecto
cd automatizador-crediflexi
```

---

### 🐍 Paso 2: Crear Entorno Virtual

Es **MUY IMPORTANTE** usar un entorno virtual para aislar las dependencias:

#### En Windows (PowerShell):
```powershell
# Crear entorno virtual
python -m venv .venv

# Activar entorno virtual
.venv\Scripts\Activate.ps1
```

**⚠️ Si te sale error de permisos en PowerShell:**
```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

#### En Mac/Linux:
```bash
# Crear entorno virtual
python3 -m venv .venv

# Activar entorno virtual
source .venv/bin/activate
```

**✅ Sabrás que está activado cuando veas `(.venv)` al inicio de tu terminal**

---

### 📦 Paso 3: Instalar Dependencias

Con el entorno virtual **activado**, instala todas las dependencias:

```bash
pip install flask flask-login flask-sqlalchemy flask-wtf pandas openpyxl werkzeug
```

**Lista completa de paquetes instalados:**
- `flask==3.1.0` - Framework web
- `flask-login==0.6.3` - Gestión de usuarios
- `flask-sqlalchemy==3.1.1` - ORM para base de datos
- `flask-wtf==1.2.1` - Formularios y CSRF
- `pandas==2.2.0` - Procesamiento de datos
- `openpyxl==3.1.2` - Lectura/escritura Excel
- `werkzeug==3.1.0` - Utilidades WSGI

**💡 Tip**: Puedes verificar las instalaciones con:
```bash
pip list
```

---

### 🗄️ Paso 4: Inicializar la Base de Datos

La base de datos se crea automáticamente la primera vez que ejecutas la aplicación. Los usuarios predefinidos son:

#### Usuario Administrador:
- **Usuario**: `admin`
- **Contraseña**: `admin123`
- **Permisos**: Generar reportes + Consultar + Eliminar

#### Usuario Consultor:
- **Usuario**: `consultor`
- **Contraseña**: `consultor123`
- **Permisos**: Solo consultar reportes

**🔐 IMPORTANTE**: Cambia estas contraseñas en producción.

---

### ▶️ Paso 5: Ejecutar la Aplicación

Con el entorno virtual **activado**:

```bash
python app.py
```

**✅ Verás algo como:**
```
 * Serving Flask app 'app'
 * Debug mode: on
 * Running on http://127.0.0.1:5000
 * Running on http://192.168.1.80:5000
```

---

### 🌐 Paso 6: Acceder a la Aplicación

Abre tu navegador web y ve a:

```
http://localhost:5000
```

O desde otra computadora en tu red local:
```
http://<TU_IP>:5000
```

**Para encontrar tu IP:**
- Windows: `ipconfig` en CMD
- Mac/Linux: `ifconfig` en Terminal

---

## 🎮 Guía de Uso

### 🔑 1. Iniciar Sesión

1. Abre `http://localhost:5000`
2. Ingresa tus credenciales:
   - Admin: `admin` / `admin123`
   - Consultor: `consultor` / `consultor123`
3. Haz clic en "Iniciar Sesión"

---

### 📊 2. Generar Reporte Individual

1. **Seleccionar archivo**
   - Haz clic en "Reporte Individual" o arrastra tu archivo Excel
   - Formato aceptado: `.xlsx` o `.xls`
   - Tamaño máximo: 16MB

2. **Opción de colaboradores** (NUEVO)
   - ✅ Marca el checkbox si quieres generar también el reporte de colaboradores
   - Esto genera 2 archivos:
     - `ReportedeAntiguedad_DDMMYYYY.xlsx` (sin códigos 001053 y 001145)
     - `ReportedeAntiguedad(Colab)_DDMMYYYY.xlsx` (solo códigos 001053 y 001145)

3. **Procesar**
   - Haz clic en "Procesar"
   - Espera a que el procesamiento termine (puede tardar 10-30 segundos)

4. **Descargar**
   - El reporte principal se descarga automáticamente
   - Si generaste el reporte de colaboradores, búscalo en el Dashboard

**📁 Archivos generados:**
- `ReportedeAntiguedad_07102025.xlsx` (nombre basado en la fecha)

---

### 👥 3. Generar Reporte Grupal

1. **Preparar 5 archivos Excel:**
   - Cobranza
   - Conformación de Grupo
   - Ahorros
   - Antigüedad de Cartera Grupal
   - Situación de Cartera

2. **Subir archivos**
   - Arrastra los 5 archivos al área de "Reporte Grupal"
   - O haz clic para seleccionarlos

3. **Procesar**
   - El sistema detecta automáticamente el tipo de cada archivo
   - Genera un reporte consolidado

---

### 📈 4. Consultar Dashboard

1. **Acceso**
   - Admin: Botón "Dashboard" en el header
   - Consultor: Acceso directo al entrar

2. **Funciones**
   - **Estadísticas**: Tarjetas con totales por tipo
   - **Filtros**: Haz clic en las tarjetas para filtrar
   - **Descarga**: Botón de descarga en cada reporte
   - **Eliminar** (solo admin): Icono de papelera

3. **Filtros disponibles:**
   - 📊 Antigüedad Individual
   - 👥 Antigüedad Grupal
   - 🕒 Recientes (últimos 7 días)

---

## ⚙️ Configuración

### 📝 config.py

Archivo principal de configuración:

#### Lista de Fraude
```python
LISTA_FRAUDE = [
    "001041", "001005", "001023", ...
]
```
**Función**: Códigos de acreditados fraudulentos que se filtran automáticamente.

#### Mapeo de Columnas
```python
COLUMN_MAPPING = {
    'dias_mora': 'Días de mora',
    'codigo_acreditado': 'Código acreditado',
    ...
}
```
**Función**: Mapeo de nombres de columnas esperados.

#### Formato Excel
```python
EXCEL_CONFIG = {
    'table_style': 'TableStyleLight1',
    'header_height': 35,
    'currency_format': '_-$* #,##0.00_-',
    ...
}
```

#### Directorio de Reportes
```python
REPORTS_FOLDER = 'static/downloads/reports'
```
**IMPORTANTE**: Los reportes se guardan aquí permanentemente.

---

## 🔧 Mantenimiento

### 🔄 Actualizar el Sistema

```bash
# Activar entorno virtual
.venv\Scripts\Activate.ps1  # Windows
source .venv/bin/activate    # Mac/Linux

# Obtener últimos cambios
git pull origin main

# Actualizar dependencias (si cambió requirements.txt)
pip install --upgrade -r requirements.txt

# Reiniciar aplicación
python app.py
```

---

### 🗑️ Limpiar Archivos Temporales

```bash
# Eliminar archivos temporales de uploads/
rm -rf uploads/*

# Limpiar cache de Python
find . -type d -name "__pycache__" -exec rm -r {} +
```

---

### 📊 Backup de Reportes

**Ubicación de reportes**: `static/downloads/reports/`

**Backup recomendado:**
```bash
# Crear copia de seguridad
cp -r static/downloads/reports/ backup/reports_$(date +%Y%m%d)/
```

---

## 🐛 Solución de Problemas

### ❌ Error: "python no se reconoce como comando"

**Solución**:
1. Reinstala Python marcando "Add to PATH"
2. O usa la ruta completa: `C:\Python313\python.exe app.py`

---

### ❌ Error: "ModuleNotFoundError: No module named 'flask'"

**Solución**:
```bash
# Verifica que el entorno virtual esté activado
# Deberías ver (.venv) en tu terminal

# Reinstala dependencias
pip install flask flask-login flask-sqlalchemy flask-wtf pandas openpyxl werkzeug
```

---

### ❌ Error: "No se puede ejecutar scripts en este sistema"

**Solución (Windows)**:
```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

---

### ❌ Error: "Failed to fetch" al procesar archivo

**Causas posibles**:
1. Archivo Excel corrupto → Verifica que se abra correctamente
2. Columnas requeridas faltantes → Revisa el formato del archivo
3. Archivo muy grande → Reduce el tamaño o incrementa `MAX_FILE_SIZE` en config.py

**Solución**:
- Revisa la consola de Python para ver el error exacto
- Verifica que el archivo tenga las columnas requeridas

---

### ❌ La columna "PAR 2" aparece duplicada

**Solución**: Este problema ya está resuelto en la versión actual. Si persiste:
1. Verifica que estés usando la última versión: `git pull`
2. Reinicia la aplicación

---

### ❌ No se encuentran códigos de colaboradores

**Causa**: Excel convierte códigos como "001053" a números (1053).

**Solución**: El sistema ya maneja ambos formatos automáticamente. Si persiste:
- Verifica que los códigos existan en el archivo original
- Revisa los logs para ver qué códigos se detectaron

---

## 📞 Soporte

### 📧 Reportar Problemas

Si encuentras un error:

1. **Revisa los logs** en la consola donde ejecutas `python app.py`
2. **Captura el error** completo (copia/pega el mensaje)
3. **Describe qué estabas haciendo** cuando ocurrió
4. **Incluye el archivo** de prueba si es posible (sin datos sensibles)

---

## 🔒 Seguridad

### ⚠️ Recomendaciones

1. **Cambiar contraseñas por defecto**
   - Edita `app/models.py` o crea un script de migración

2. **No exponer a Internet sin HTTPS**
   - Usa solo en red local
   - O configura un proxy inverso (nginx) con SSL

3. **Backup regular**
   - Base de datos: `instance/crediflexi.db`
   - Reportes: `static/downloads/reports/`

4. **Actualizar dependencias**
   ```bash
   pip list --outdated
   pip install --upgrade <paquete>
   ```

---

## 📄 Licencia

Este proyecto es software propietario desarrollado para CREDIFLEXI. Todos los derechos reservados.

---

## 👥 Créditos

**Desarrollado para**: CREDIFLEXI  
**Tecnología**: Python, Flask, Pandas, OpenPyXL  
**Versión**: 2.0 (Octubre 2025)

---

## 🎯 Próximas Funcionalidades (Roadmap)

- [ ] Exportar reportes a PDF
- [ ] Notificaciones por email cuando se genera un reporte
- [ ] Gráficas y visualizaciones en el dashboard
- [ ] API REST para integración con otros sistemas
- [ ] Programación de generación automática de reportes

---

**CREDIFLEXI** - Soluciones Profesionales para Gestión de Cartera Crediticia

