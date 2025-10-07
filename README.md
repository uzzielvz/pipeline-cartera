# CREDIFLEXI - Sistema de Automatización de Reportes

## Descripción General

CREDIFLEXI es una aplicación web profesional diseñada para automatizar el procesamiento de reportes de antigüedad de cartera crediticia. El sistema proporciona procesamiento automatizado de datos, filtrado de fraudes, integración de geolocalización y generación de reportes Excel profesionales con formato avanzado y estilos condicionales.

---

## Características Principales

### Funcionalidad Core

#### Reporte de Antigüedad Individual
- Procesamiento automatizado de archivos Excel con datos de cartera crediticia
- Filtrado automático de códigos fraudulentos usando listas predefinidas
- Generación de enlaces de geolocalización en Google Maps para verificación de direcciones
- Cálculo automático de PAR (Portfolio at Risk) basado en días de mora
- Opción para generar reporte separado de colaboradores (códigos 001053 y 001145)

#### Reporte de Antigüedad Grupal
- Integración de 5 archivos Excel para generar reporte consolidado
- Detección automática del tipo de cada archivo por contenido
- Procesamiento unificado de datos de múltiples fuentes

#### Sistema de Usuarios y Autenticación
- **Administrador/Generador**: Puede generar reportes y acceder al dashboard completo
- **Consultor**: Solo puede consultar reportes existentes en el dashboard
- Autenticación segura con Flask-Login
- Gestión de permisos por rol de usuario

#### Dashboard Interactivo
- Visualización de estadísticas de reportes generados
- Filtros dinámicos por tipo de reporte (Individual, Grupal, Colaboradores)
- Historial completo de reportes con información detallada (fecha, usuario, tamaño de archivo)
- Descarga de reportes anteriores
- Eliminación de reportes (solo para administradores)

### Generación de Reportes

Cada reporte individual genera un archivo Excel con las siguientes hojas:

1. **Informe Completo** (fecha del día)
   - Dataset completo con todos los registros procesados
   - Código acreditado como primera columna para fácil identificación
   - Enlaces de geolocalización clickeables para cada dirección
   - Formato profesional con tablas y estilos condicionales

2. **Mora**
   - Registros filtrados con 1 o más días de mora
   - Columnas adicionales para seguimiento operativo:
     - **Seguimiento Call Center** (4 columnas con fondo verde)
     - **Gestión de Cobranza en Campo** (5 columnas con fondo azul)

3. **Cuentas con saldo vencido**
   - Registros que cumplen: saldo vencido mayor o igual a 1 y días de mora menores o iguales a 0
   - Filtrado especializado para identificación de casos especiales

4. **Liquidación anticipada**
   - Hoja interactiva con calculadora de liquidación
   - Fórmulas VLOOKUP configuradas para búsqueda automática
   - Usuario ingresa código acreditado y obtiene cálculo de montos
   - Tabla de referencia con datos completos del reporte

5. **Hojas por Coordinación**
   - Una hoja independiente por cada coordinación/región
   - Datos filtrados y organizados por área geográfica
   - Facilita la distribución de trabajo por equipos

### Capacidades de Procesamiento de Datos

- **Limpieza de Datos**: Estandariza números telefónicos y formatos de datos
- **Mapeo Inteligente de Columnas**: Detección automática de columnas con nombres variables
- **Prevención de Duplicados**: Manejo automático de columnas duplicadas (ej. PAR)
- **Validación de Integridad**: Validación exhaustiva de datos durante todo el proceso
- **Manejo de Errores**: Sistema robusto de manejo de excepciones con logging detallado

---

## Arquitectura Técnica

### Stack Tecnológico

- **Backend**: Python 3.13 con Flask 3.1.0
- **Procesamiento de Datos**: Pandas 2.2.0 para manipulación y análisis
- **Generación Excel**: OpenPyXL 3.1.2 para creación avanzada de archivos Excel
- **Base de Datos**: SQLite con SQLAlchemy 3.1.1
- **Autenticación**: Flask-Login 0.6.3
- **Formularios y Seguridad**: Flask-WTF 1.2.1 con protección CSRF
- **Utilidades**: Werkzeug 3.1.0
- **Frontend**: HTML5, CSS3, JavaScript (Vanilla)

### Estructura del Proyecto

```
automatizador-crediflexi/
├── app/
│   ├── __init__.py          # Inicialización del paquete
│   ├── reportes.py          # Lógica de procesamiento de reportes
│   ├── auth.py              # Sistema de autenticación y login
│   ├── consultor.py         # Dashboard y consultas de reportes
│   └── models.py            # Modelos de base de datos (User, ReportHistory)
├── static/
│   ├── css/
│   │   └── style.css        # Estilos de la aplicación
│   ├── downloads/
│   │   └── reports/         # Almacenamiento permanente de reportes generados
│   └── js/                  # JavaScript del cliente
├── templates/
│   ├── base.html            # Template base con estructura común
│   ├── index.html           # Página principal para generación de reportes
│   ├── auth/
│   │   └── login.html       # Página de inicio de sesión
│   ├── consultor/
│   │   ├── dashboard.html   # Dashboard principal de reportes
│   │   ├── reports.html     # Lista completa de reportes
│   │   └── report_detail.html
│   └── errors/
│       └── unauthorized.html # Página de acceso no autorizado
├── uploads/                 # Almacenamiento temporal de archivos subidos
├── instance/
│   └── crediflexi.db        # Base de datos SQLite
├── app.py                   # Punto de entrada de la aplicación Flask
├── config.py                # Configuración general del sistema
├── .gitignore               # Archivos y directorios ignorados por Git
├── requirements.txt         # Dependencias Python del proyecto
└── README.md                # Este archivo
```

---

## Instalación Local

### Prerrequisitos

Antes de comenzar, asegúrese de tener instalado:

1. **Python 3.13 o superior**
   - Descargar desde: https://www.python.org/downloads/
   - Durante la instalación en Windows, marcar la opción "Add Python to PATH"

2. **Git** (para clonar el repositorio)
   - Descargar desde: https://git-scm.com/downloads

3. **Editor de código** (recomendado)
   - Visual Studio Code: https://code.visualstudio.com/
   - PyCharm, Sublime Text, o cualquier editor de su preferencia

---

### Paso 1: Clonar el Repositorio

Abra una terminal (PowerShell en Windows, Terminal en Mac/Linux) y ejecute:

```bash
# Clonar el repositorio
git clone <URL_DEL_REPOSITORIO>

# Entrar al directorio del proyecto
cd automatizador-crediflexi
```

---

### Paso 2: Crear y Activar Entorno Virtual

Es importante usar un entorno virtual para aislar las dependencias del proyecto:

#### En Windows (PowerShell):
```powershell
# Crear entorno virtual
python -m venv .venv

# Activar entorno virtual
.venv\Scripts\Activate.ps1
```

**Nota**: Si aparece un error de permisos en PowerShell, ejecute:
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

**Verificación**: Cuando el entorno esté activado, verá `(.venv)` al inicio de la línea de comandos.

---

### Paso 3: Instalar Dependencias

Con el entorno virtual activado, instale todas las dependencias necesarias:

```bash
pip install flask flask-login flask-sqlalchemy flask-wtf pandas openpyxl werkzeug
```

**Lista completa de paquetes y versiones:**
- `flask==3.1.0` - Framework web principal
- `flask-login==0.6.3` - Gestión de sesiones de usuario
- `flask-sqlalchemy==3.1.1` - ORM para base de datos
- `flask-wtf==1.2.1` - Formularios web y protección CSRF
- `pandas==2.2.0` - Procesamiento y análisis de datos
- `openpyxl==3.1.2` - Lectura y escritura de archivos Excel
- `werkzeug==3.1.0` - Utilidades WSGI y seguridad

**Verificar instalación:**
```bash
pip list
```

---

### Paso 4: Inicializar la Base de Datos

La base de datos SQLite se crea automáticamente la primera vez que ejecuta la aplicación.

**Usuarios predefinidos:**

| Tipo | Usuario | Contraseña | Permisos |
|------|---------|-----------|----------|
| Administrador | `admin` | `admin123` | Generar reportes, consultar dashboard, eliminar reportes |
| Administrador | `juan.carlos` | `juancarlos123` | Generar reportes, consultar dashboard, eliminar reportes |
| Consultor | `consultor` | `consultor123` | Solo consultar reportes existentes |

**IMPORTANTE**: En un entorno de producción, cambie estas contraseñas por defecto.

---

### Paso 5: Ejecutar la Aplicación

Con el entorno virtual activado, ejecute:

```bash
python app.py
```

**Salida esperada:**
```
 * Serving Flask app 'app'
 * Debug mode: on
WARNING: This is a development server. Do not use it in a production deployment.
 * Running on http://127.0.0.1:5000
 * Running on http://192.168.x.x:5000
Press CTRL+C to quit
```

---

### Paso 6: Acceder a la Aplicación

#### Acceso Local:
Abra su navegador web y navegue a:
```
http://localhost:5000
```

#### Acceso desde Red Local:
Para acceder desde otra computadora en la misma red:
```
http://<IP_DE_SU_COMPUTADORA>:5000
```

**Para encontrar su dirección IP:**
- Windows: Ejecute `ipconfig` en CMD y busque "Dirección IPv4"
- Mac/Linux: Ejecute `ifconfig` o `ip addr` en Terminal

---

## Guía de Uso

### 1. Inicio de Sesión

1. Abra la aplicación en su navegador: `http://localhost:5000`
2. Ingrese sus credenciales:
   - **Admin**: `admin` / `admin123`
   - **Consultor**: `consultor` / `consultor123`
3. Haga clic en "Iniciar Sesión"

---

### 2. Generar Reporte Individual

**Paso a paso:**

1. **Preparar archivo**
   - Formato aceptado: `.xlsx` o `.xls`
   - Tamaño máximo: 16MB
   - Debe contener las columnas requeridas (código acreditado, días de mora, coordinación, etc.)

2. **Subir archivo**
   - Opción A: Haga clic en el área "Reporte Individual" y seleccione el archivo
   - Opción B: Arrastre y suelte el archivo en el área designada

3. **Configurar opciones**
   - Marque el checkbox "Generar también reporte de colaboradores" si desea generar dos reportes:
     - **Reporte principal**: Excluye códigos 001053 y 001145
     - **Reporte de colaboradores**: Solo incluye códigos 001053 y 001145

4. **Procesar**
   - Haga clic en el botón "Procesar"
   - El sistema procesará el archivo (tiempo estimado: 10-30 segundos dependiendo del tamaño)

5. **Descargar**
   - El reporte principal se descarga automáticamente
   - El reporte de colaboradores (si fue solicitado) estará disponible en el Dashboard

**Archivos generados:**
- Formato de nombre: `ReportedeAntiguedad_DDMMYYYY.xlsx`
- Si generó reporte de colaboradores: `ReportedeAntiguedad(Colab)_DDMMYYYY.xlsx`

---

### 3. Generar Reporte Grupal

**Requisitos:**

Debe preparar 5 archivos Excel correspondientes a:
1. Reporte de Cobranza
2. Conformación de Grupo
3. Reporte de Ahorros
4. Antigüedad de Cartera Grupal
5. Situación de Cartera

**Proceso:**

1. **Subir archivos**
   - Arrastre los 5 archivos al área de "Reporte Grupal"
   - O haga clic para seleccionarlos desde el explorador de archivos

2. **Detección automática**
   - El sistema detecta automáticamente el tipo de cada archivo basándose en su contenido
   - No es necesario que los archivos tengan nombres específicos

3. **Procesar**
   - Haga clic en "Procesar"
   - El sistema genera un reporte consolidado integrando los 5 archivos

4. **Descargar**
   - El reporte grupal se descarga automáticamente

---

### 4. Consultar Dashboard

**Acceso:**
- **Administradores**: Botón "Dashboard" en el encabezado de la página
- **Consultores**: Acceso directo al iniciar sesión

**Funcionalidades:**

1. **Estadísticas Generales**
   - Tarjetas informativas muestran totales por tipo de reporte
   - Total de reportes generados
   - Reportes de antigüedad individual
   - Reportes de antigüedad grupal
   - Reportes recientes (últimos 7 días)

2. **Filtros Dinámicos**
   - Haga clic en las tarjetas de estadísticas para filtrar la lista
   - Filtre por tipo: Individual, Grupal, Colaboradores
   - Filtre por fecha: Recientes (últimos 7 días)

3. **Gestión de Reportes**
   - **Descargar**: Botón de descarga en cada entrada de reporte
   - **Eliminar** (solo admin): Icono de papelera para borrar reportes
   - **Información detallada**: Fecha de generación, usuario, tamaño de archivo

4. **Búsqueda y Navegación**
   - Lista ordenada por fecha (más recientes primero)
   - Información de tipo de reporte y usuario generador

---

## Configuración

### Archivo config.py

Este archivo contiene todas las configuraciones principales del sistema:

#### Lista de Fraude
```python
LISTA_FRAUDE = [
    "001041", "001005", "001023", "001024", "001025",
    # ... más códigos
]
```
**Propósito**: Códigos de acreditados que se filtran automáticamente por ser fraudulentos.

#### Mapeo de Columnas
```python
COLUMN_MAPPING = {
    'dias_mora': 'Días de mora',
    'codigo_acreditado': 'Código acreditado',
    'nombre_cliente': 'Nombre del cliente',
    # ... más mapeos
}
```
**Propósito**: Permite que el sistema reconozca columnas aunque tengan nombres ligeramente diferentes.

#### Configuración de Excel
```python
EXCEL_CONFIG = {
    'table_style': 'TableStyleLight1',
    'header_height': 35,
    'currency_format': '_-$* #,##0.00_-',
    'percentage_format': '0.00%',
    # ... más configuraciones
}
```
**Propósito**: Define estilos, formatos y apariencia de los reportes Excel generados.

#### Directorios de Almacenamiento
```python
UPLOAD_FOLDER = 'uploads'           # Archivos temporales
REPORTS_FOLDER = 'static/downloads/reports'  # Reportes permanentes
```
**IMPORTANTE**: Los reportes se guardan permanentemente en `static/downloads/reports/`.

---

## Mantenimiento

### Actualizar el Sistema

```bash
# Activar entorno virtual
.venv\Scripts\Activate.ps1  # Windows
source .venv/bin/activate    # Mac/Linux

# Obtener últimos cambios del repositorio
git pull origin main

# Actualizar dependencias (si requirements.txt cambió)
pip install --upgrade -r requirements.txt

# Reiniciar aplicación
python app.py
```

---

### Limpiar Archivos Temporales

Los archivos en la carpeta `uploads/` son temporales y pueden eliminarse periódicamente:

```bash
# Windows PowerShell
Remove-Item -Path uploads\* -Force

# Mac/Linux
rm -rf uploads/*
```

**Nota**: No elimine archivos de `static/downloads/reports/` ya que contiene los reportes permanentes.

---

### Backup de Datos

**Archivos críticos para respaldo:**

1. **Base de datos**: `instance/crediflexi.db`
   - Contiene usuarios y historial de reportes

2. **Reportes generados**: `static/downloads/reports/`
   - Contiene todos los reportes Excel generados

**Ejemplo de backup:**
```bash
# Crear directorio de backup
mkdir backup_crediflexi_$(date +%Y%m%d)

# Copiar base de datos
cp instance/crediflexi.db backup_crediflexi_$(date +%Y%m%d)/

# Copiar reportes
cp -r static/downloads/reports/ backup_crediflexi_$(date +%Y%m%d)/
```

---

## Solución de Problemas

### Error: "python no se reconoce como comando"

**Causa**: Python no está en el PATH del sistema.

**Solución**:
1. Reinstale Python marcando la opción "Add Python to PATH"
2. O use la ruta completa: `C:\Python313\python.exe app.py`

---

### Error: "ModuleNotFoundError: No module named 'flask'"

**Causa**: Las dependencias no están instaladas o el entorno virtual no está activado.

**Solución**:
```bash
# Verificar que el entorno virtual esté activado
# Debe ver (.venv) al inicio de la línea de comandos

# Reinstalar dependencias
pip install flask flask-login flask-sqlalchemy flask-wtf pandas openpyxl werkzeug
```

---

### Error: "No se puede ejecutar scripts en este sistema" (Windows)

**Causa**: Política de ejecución de PowerShell restrictiva.

**Solución**:
```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

---

### Error: "Failed to fetch" al procesar archivo

**Causas posibles**:
1. Archivo Excel corrupto o con formato incorrecto
2. Columnas requeridas faltantes en el archivo
3. Archivo demasiado grande

**Solución**:
1. Verifique que el archivo se abra correctamente en Excel
2. Revise la consola de Python para ver el mensaje de error específico
3. Asegúrese de que el archivo contenga todas las columnas requeridas
4. Si el archivo es muy grande, considere dividirlo o aumentar `MAX_FILE_SIZE` en `config.py`

---

### La columna "PAR" aparece duplicada

**Causa**: Versión desactualizada del código.

**Solución**:
```bash
git pull origin main
python app.py
```

Este problema está resuelto en la versión actual.

---

### No se encuentran códigos de colaboradores (001053, 001145)

**Causa**: Excel puede convertir estos códigos de texto a números (1053, 1145).

**Solución**: El sistema ya maneja ambos formatos automáticamente. Si persiste:
1. Verifique que los códigos existan en el archivo original
2. Revise los logs en la consola para ver qué códigos se detectaron

---

## Seguridad

### Recomendaciones de Seguridad

1. **Cambiar contraseñas por defecto**
   - Las contraseñas `admin123`, `consultor123`, etc., deben cambiarse en producción
   - Modifique el archivo `app/models.py` o cree un script de migración

2. **No exponer a Internet sin HTTPS**
   - Este servidor Flask es de desarrollo
   - Para producción, use un servidor WSGI (Gunicorn, uWSGI) detrás de un proxy inverso (nginx, Apache) con SSL/TLS

3. **Backup regular**
   - Base de datos: `instance/crediflexi.db`
   - Reportes: `static/downloads/reports/`
   - Configure backups automáticos diarios o semanales

4. **Actualizar dependencias**
   ```bash
   # Ver paquetes desactualizados
   pip list --outdated
   
   # Actualizar paquete específico
   pip install --upgrade <paquete>
   ```

5. **Protección de archivos sensibles**
   - El archivo `.gitignore` ya está configurado para no subir:
     - Base de datos (`instance/*.db`)
     - Archivos subidos (`uploads/`)
     - Reportes generados (`static/downloads/`)
     - Entorno virtual (`.venv/`)

---

## Control de Versiones

### Archivos Ignorados por Git

El archivo `.gitignore` está configurado para excluir:

```
# Python
__pycache__/
*.pyc

# Datos y reportes
static/downloads/
uploads/
instance/*.db

# Entorno virtual
.venv/

# IDE
.vscode/
.idea/
```

### Hacer Commit de Cambios

```bash
# Ver estado actual
git status

# Agregar archivos modificados
git add <archivo>

# Hacer commit
git commit -m "Descripción de cambios"

# Subir a repositorio remoto
git push origin main
```

---

## Soporte Técnico

### Reportar Problemas

Si encuentra un error o bug:

1. **Revisar los logs**
   - Consulte la salida de la consola donde ejecuta `python app.py`
   - Los mensajes de error contienen información valiosa

2. **Capturar información del error**
   - Copie el mensaje de error completo
   - Incluya el stack trace si está disponible

3. **Describir el contexto**
   - ¿Qué estaba haciendo cuando ocurrió el error?
   - ¿Qué archivo estaba procesando?
   - ¿Es reproducible el error?

4. **Información del sistema**
   - Versión de Python: `python --version`
   - Sistema operativo y versión
   - Versiones de dependencias: `pip list`

---

## Licencia

Este proyecto es software propietario desarrollado para CREDIFLEXI. Todos los derechos reservados.

El uso, copia, modificación y distribución de este software está restringido a personal autorizado de CREDIFLEXI.

---

## Información del Proyecto

**Desarrollado para**: CREDIFLEXI  
**Tecnología Principal**: Python, Flask, Pandas, OpenPyXL  
**Versión**: 2.0  
**Última actualización**: Octubre 2025

---

## Contacto

Para soporte técnico, consultas o reportar problemas, contacte al equipo de desarrollo o al administrador del sistema.

---

**CREDIFLEXI**
