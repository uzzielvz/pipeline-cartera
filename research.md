# CREDIFLEXI / automatizador-crediflexi — Research / Contexto

> Documento de referencia vivo. Última actualización: 2026-06-04 (commit `fe8d8e0`).
> Docs complementarios (no borrar): `research_cambios.md` (detalle de iteraciones del nuevo formato), `FINAL_RESEARCH.md` (diff sistema actual vs `FINAL TARGET.xlsx`), `README.md` (manual de uso/instalación).

## Qué es y para qué
App web interna de CREDIFLEXI que **automatiza la generación de reportes de antigüedad de cartera crediticia** en Excel. Un usuario administrador sube el Excel crudo exportado del sistema administrador de créditos (Reporte de Antigüedad de Cartera) y el sistema devuelve un `.xlsx` con múltiples hojas, columnas calculadas, pivots, formato condicional y filtros de negocio (fraude, recuperadores, situación de crédito). Consultores solo visualizan/descargan reportes históricos.

Hay **dos flujos**:
- **Individual** — IMPLEMENTADO y maduro. Es el flujo vivo en uso. Toda la lógica real está aquí.
- **Grupal** — STUB. El endpoint existe pero genera un Excel dummy de 1 hoja (TODO sin implementar). Ver §Estado actual.

## Stack y arquitectura
- **Lenguaje**: Python 3.13 (evidencia: `__pycache__/*.cpython-313.pyc`).
- **Web**: Flask 3.1 (factory `create_app` en `app.py`, blueprints).
- **Datos**: pandas (manipulación de DataFrames).
- **Excel**: openpyxl (escritura celda a celda, tablas, formato, hipervínculos; los pivots vienen de una plantilla `.xlsx`, openpyxl no los genera).
- **DB**: SQLite vía Flask-SQLAlchemy (`instance/crediflexi.db`).
- **Auth**: Flask-Login + Flask-WTF (CSRF). Roles ADMIN / CONSULTOR.
- **Frontend**: Jinja2 + HTML/CSS/JS vanilla.
- ⚠️ **No existe `requirements.txt`** en el repo (el README lo menciona pero el archivo no está). Las versiones reales viven en `.venv`. Suposición: pandas 2.x / openpyxl 3.1.x (verificado en sesiones previas, ver `research.md` histórico).

Diagrama de alto nivel:
```
Navegador (templates/index.html)
   │  POST multipart
   ▼
Flask blueprint /reportes  (app/reportes.py)
   ├─ /procesar_antiguedad        → procesar_reporte_antiguedad()  [FLUJO VIVO]
   └─ /procesar_antiguedad_grupal → procesar_antiguedad_grupal()   [STUB/dummy]
   │
   ▼
pandas (limpieza, filtros, columnas calculadas)
   │
   ▼
openpyxl: copia PLANTIILA2.xlsx → rellena hojas → guarda en uploads/
   │
   ▼
move a static/downloads/reports/  + registro en ReportHistory (SQLite)
   │
   ▼
send_file (descarga) ; Consultor dashboard lista/descarga históricos
```

## Estructura del repo
```
app.py                      # Entry point Flask (create_app); corre en host 0.0.0.0 port 5001
config.py                   # Constantes: COLUMN_MAPPING, LISTA_FRAUDE, CODIGOS_RECUPERADOR_EXCLUIR,
                            #   PERIODICIDAD_A_DIAS, EXCEL_CONFIG, COLORS, ADDITIONAL_COLUMNS, roles/permisos, SECRET_KEY
parches.json                # Parches manuales (sobrescribe campos por codigo+ciclo). Ver §Modelo/Convenciones
app/__init__.py             # Paquete (vacío)
app/reportes.py             # ★ NÚCLEO: ~3150 líneas. Todo el pipeline individual + grupal stub + rutas
app/auth.py                 # Login/logout, init_auth, seed de usuarios, require_permission
app/consultor.py            # Dashboard, lista y descarga de reportes históricos
app/models.py               # SQLAlchemy: User, ReportHistory
templates/                  # index.html (upload), auth/login, consultor/*, errors/unauthorized, base.html
static/css/style.css        # Estilos
static/downloads/reports/   # Reportes generados (permanentes, .gitignored)
uploads/                    # Temporal de subida + salida intermedia (.gitignored)
instance/crediflexi.db      # SQLite (.gitignored)
PLANTIILA2.xlsx             # ★ Plantilla con pivots/tablas que el código copia y rellena (flujo vivo)
PLANTIILA2_nueva.xlsx.bak   # Backup de plantilla alternativa (rama nueva-plantilla; no usada en main)
FINAL TARGET.xlsx           # Excel objetivo de referencia para el nuevo formato
test_abril2026_input.xlsx   # Input crudo de ejemplo (1 hoja, headers con \n, 342 filas)
ReportedeAntiguedad_*.xlsx  # Outputs/targets de ejemplo para comparación
research.md / plan.md       # Docs vivos (este y la hoja de ruta)
research_cambios.md / FINAL_RESEARCH.md  # Docs de referencia de las iteraciones de formato
```

## Modelo de datos
SQLite, dos tablas (`app/models.py`):
- **users**: `id, username (unique), email (unique), password_hash, role ('ADMIN'|'CONSULTOR'), is_active, created_at, last_login`. Métodos: `check_password`, `has_permission`, `is_admin`, `is_consultor`.
- **report_history**: `id, user_id (FK users.id), report_type ('individual'|'grupal'), filename, file_path, created_at, file_size`. Relación `User.reports` (1‑N).

`parches.json` (no es DB, es config de datos): lista `parches[]`; cada parche = `{id, descripcion, activo, campo, nuevo_valor, registros[{codigo, ciclo, nombre}]}`. Sobrescribe `campo`=`nuevo_valor` para filas que casen por `Código acreditado`(6 díg) + `Ciclo`(2 díg). `nombre` es solo referencia humana.

## Integraciones externas
- **Google Maps**: se generan URLs de geolocalización (`https://www.google.com/maps?q=lat,lng`) a partir de la columna de geolocalización. **No usa API key** — solo construcción de URL (`generar_link_google_maps`, `add_geolocation_links`).
- No hay otras APIs ni servicios externos. No hay credenciales de terceros.
- ⚠️ `SECRET_KEY` de Flask está **hardcodeada** en `config.py` (debe moverse a variable de entorno — deuda de seguridad). Usuarios semilla con contraseñas por defecto se crean en el arranque (documentadas en README; cambiar en producción).

## INPUT → TRANSFORMACIONES → OUTPUT (flujo individual, el vivo)

### INPUT
- Archivo `.xlsx`/`.xls` subido a `/reportes/procesar_antiguedad` (campo `archivo`). Máx 16MB (`MAX_FILE_SIZE`).
- Es el **Reporte de Antigüedad de Cartera** crudo: 1 hoja, `header=0`, headers con saltos de línea `\n`, ~63 columnas. Columnas clave: `Código acreditado`, `Código promotor`, `Código recuperador`, `Ciclo`, `Inicio ciclo`, `Días de mora`, `Coordinación`, `Situación crédito`, `Saldo vencido`, `Saldo total`, `Saldo riesgo total`, `Último pago`, `Periodicidad`, `Geolocalización domicilio`, medios de comunicación y referencias. Listado completo de columnas en `research_cambios.md`.

### TRANSFORMACIONES — `procesar_reporte_antiguedad()` (`app/reportes.py:1259`)
Orden exacto de pasos (con nº de línea aproximado):
1. **PASO 1 (1270)** `read_excel(header=0)` + `clean_dataframe_columns` → quita `\n` de nombres de columna (`Situación\ncrédito` → `Situación crédito`).
2. **PASO 1.1 (1301)** `standardize_codes` → `Código acreditado/promotor/recuperador` a 6 dígitos con zfill.
3. **PASO 1.1b (1306)** `aplicar_parches` → aplica `parches.json` (1er parche: reasigna `Coordinación`=`Oficina Central` a 18 acreditados por codigo+ciclo).
4. **PASO 2 (1309)** filtra `LISTA_FRAUDE` (códigos de acreditado fraudulentos) → `df_filtrado`.
5. **PASO 2.1 (1324)** separa subset **RECUPERADOR_000124**: filas con `Código recuperador` en `CODIGOS_RECUPERADOR_EXCLUIR` salen de `df_filtrado` a `df_recup_000124_raw`.
6. **PASO 2.2 (1341)** separa subset **Autorizado por Cartera**: filas con `Situación crédito` ≠ `Entregado` (case-insensitive, ignora vacíos) salen de `df_filtrado` a `df_autcart_raw`. (RECUPERADOR tiene prioridad por orden.)
7. **PASO 1.2 (1358)** sobre `df_filtrado`: `clean_phone_numbers`, formato `Ciclo` a 2 díg, `add_geolocation_links`.
8. **PASO 1.3 (1378)** construye `df_completo`: orden por `Días de mora` desc, `add_par_column`, inserta `Link de Geolocalización` tras geolocalización, reordena `Código acreditado` primero, elimina duplicadas.
9. **Pipelines de subsets (1412, 1437)** RECUPERADOR_000124 y Autorizado por Cartera pasan por el mismo pipeline (teléfonos, ciclo, geo, orden, PAR, concepto depósito, riesgo/mora, días último pago/alerta).
10. **PASO 3 (1462)** `df_ordenado` orden por mora + `add_par_column`.
11. **PASO 4 (1483)** `df_mora` = registros con `Días de mora` ≥ 1.
12. **PASO 4.1 (1497)** `df_saldo_vencido` = `Saldo vencido` ≥ 1 AND `Días de mora` ≤ 0.
13. **PASO 6 (1545)** generación Excel (ver OUTPUT).

Columnas calculadas (funciones que las producen):
- `PAR` — `add_par_column` / `asignar_rango_mora`: categorías por días de mora (0, 7, 15, 30, 60, 90, Mayor_90, Mayor_180).
- `Concepto Depósito` — `agregar_columna_concepto_deposito`.
- `% MORA`, `Saldo riesgo capital/total` — `agregar_columnas_riesgo_y_mora`.
- `Días desde último pago`, `Alerta` — `agregar_columnas_dias_ultimo_pago_y_alerta` (usa `PERIODICIDAD_A_DIAS`).
- `Cuotas sin pagar`, `Saldo_Riesgo_total`, `Combinado` — `agregar_columnas_nuevas` (solo R_Completo).
- `Suma` (col 75) — 1 si `Días de mora` ∈ [1,30] else 0.

### OUTPUT — Excel con plantilla (`PLANTIILA2.xlsx`)
Si existe la plantilla (caso normal): se **copia** y se rellenan/crean hojas. Nombre de salida `ReportedeAntigüedad_DDMMYYYY.xlsx` donde la fecha = día anterior (o viernes si hoy es lunes).

Hojas resultantes (orden final, ITERACIÓN 14 en `:1837`):
1. **R_Completo** — 74 columnas + `Suma`. Datos desde fila 3, headers de la plantilla (fila 2). Formato condicional en mora, % mora, alerta (rojo), moneda/fecha, hipervínculos geo, rango de tabla actualizado.
2. **[DDMMYYYY]** — copia de R_Completo (hoja con nombre de fecha).
3. **Marzo2026** — histórico: todos los registros con `Inicio ciclo` < 2026-04-01 (`:1777`).
4. **Abril2026** — registros con `Inicio ciclo` ≥ 2026-04-01 (`:1705`).
5. **X_Coordinación** — pivots (vienen de la plantilla; caches apuntan a `Marzo2026`/`Abril2026` por nombre fijo, ver FINAL_RESEARCH).
6. **X_Recuperador** — pivots de la plantilla.
7. **RECUPERADOR_000124** — subset por código recuperador (`:2541`).
8. **Autorizado por Cartera** — subset `Situación crédito` ≠ Entregado (`:2565`).
9. **Mora** — registros con mora ≥ 1, columnas reordenadas (empieza `Nom. región`, ITERACIÓN 6 `:2614`), + columnas operativas (Seguimiento Call Center / Gestión Cobranza Campo, `ADDITIONAL_COLUMNS`).
10. **Cuentas con saldo vencido** — `:2650`.
11. **Liquidación anticipada** — calculadora con fórmulas VLOOKUP a R_Completo (`:2706`).

Sin plantilla (fallback): el código crea X_Coordinación/X_Recuperador con pandas (`crear_hoja_x_coordinacion`/`crear_hoja_x_recuperador`, `:603/:792`) y escribe vía `ExcelWriter`.

## Comandos clave
- **Dev/run**: `python app.py` → http://localhost:5001 (⚠️ README dice 5000, código dice 5001).
- **Entorno**: `python -m venv .venv` + activar; deps (no hay requirements.txt): `pip install flask flask-login flask-sqlalchemy flask-wtf pandas openpyxl werkzeug`.
- **DB**: se crea sola al arrancar; usuarios semilla se siembran en `init_auth` (`app/auth.py`).
- **Tests**: no hay. La validación es manual: subir un input y comparar el Excel contra `FINAL TARGET.xlsx`.
- **Build/deploy**: no hay pipeline. Es servidor de desarrollo Flask (no apto producción sin WSGI+proxy).
- **Python directo del venv** (útil para inspección): `.venv/Scripts/python -c "..."` (Windows, shell bash).

## Decisiones tomadas y por qué
- **Plantilla Excel con pivots precargados** en vez de generar pivots por código: openpyxl no soporta crear pivots; se copia `PLANTIILA2.xlsx` y solo se rellenan datos. Los pivot caches están atados a nombres fijos `Marzo2026`/`Abril2026`.
- **Filtros de negocio como separación de subsets** (no solo ocultar): fraude se elimina; RECUPERADOR_000124 y Autorizado por Cartera se sacan del informe principal y van a su propia hoja. Patrón replicable (los dos últimos comparten pipeline idéntico).
- **`parches.json` externo** para correcciones manuales puntuales (reasignar campos por codigo+ciclo) sin tocar código — pensado para crecer con más parches.
- **Fecha del reporte = día hábil anterior** (viernes si lunes): refleja el corte operativo.
- **Hojas por coordinación eliminadas** (iter 1, commit `92f541a`) — ya no se generan.
- **Checkbox/reporte de colaboradores (001053/001145) retirado** (commit `3a84ecb`) aunque el README aún lo describe.

## Convenciones del proyecto
- Comentarios de pipeline marcados como `# --- PASO X ---` y `# --- ITERACIÓN N ---` (orden cronológico de features).
- Funciones auxiliares de transformación nombradas `agregar_columna(s)_*`, `aplicar_formato_*`, `crear_hoja_*`.
- Códigos siempre normalizados a 6 dígitos (`zfill(6)`), Ciclo a 2 (`zfill(2)`).
- Logging con emojis (✅ éxito, ⚠️ warning, 🔍 filtro, 📋 hoja, 🩹 parche).
- Constantes de negocio centralizadas en `config.py`.
- Commits en inglés con prefijo `feat:`/`fix:`/`chore:` y a menudo sufijo `(iter N)`.
- **Política de commits**: NUNCA incluir `Co-Authored-By: Claude` (política de la empresa).

## Estado actual
**Terminado / estable**:
- Flujo individual completo: carga, limpieza, fraude, parches, subsets RECUPERADOR_000124 y Autorizado por Cartera, columnas calculadas, 11 hojas con formato, plantilla con pivots, descarga e historial.
- Auth con roles, dashboard de consultor, historial en SQLite.
- Sistema de `parches.json` (commit `dcc07bb`) y hoja Autorizado por Cartera (`fe8d8e0`).

**A medias / pendiente**:
- **Flujo grupal**: STUB. `procesar_antiguedad_grupal` (`:3141`) detecta 5 tipos de archivo (`detectar_tipo_archivo`, `:3111`) pero genera un Excel dummy de 1 hoja. TODO explícito sin lógica de consolidación. Faltan: spec de columnas/hojas de salida, cruces entre los 5 archivos, ejemplos de input. (Detalle en `research.md` histórico / FINAL_RESEARCH.)
- **Criterio configurable de Autorizado por Cartera**: hoy hardcodeado a `≠ Entregado`. Pendiente decisión del dueño (solo `Autorizado por cartera`/`tesorería`, o lista configurable en `config.py`).

**Fase**: el flujo individual está en pulido iterativo para igualar `FINAL TARGET.xlsx` + se agregan filtros operativos. El grupal no ha comenzado.

## Riesgos / deuda técnica conocida
- **Hardcodes mensuales**: `Marzo2026`, `Abril2026` y la fecha de corte `2026-04-01` están hardcodeados en código y en los pivot caches de la plantilla → **hay que editarlos cada mes** o el reporte queda mal. Riesgo recurrente #1.
- **`SECRET_KEY` y contraseñas por defecto** en código → deben ir a variables de entorno.
- **No hay `requirements.txt`** → onboarding frágil, versiones no fijadas.
- **No hay tests** → toda validación es manual contra `FINAL TARGET.xlsx`.
- **`df_r_completo.iloc[:, :74]`** recorte por posición fija → frágil si cambian columnas del input.
- **Nombre de columna `Situación crédito`** y valor `Entregado` hardcodeados en el filtro nuevo.
- **Plantilla path** `PLANTIILA2.xlsx` hardcodeado; existe `PLANTIILA2_nueva.xlsx.bak` de otra rama (posible confusión de cuál es la buena).
- **Discrepancia de puerto** (README 5000 vs código 5001) y de features (README describe colaboradores y hojas por coordinación ya retirados).
- **Grupal**: errores silenciosos en detección de tipos (devuelve 400/500 sin detalle), sin esquema de columnas.
- **Encoding**: nombres de hoja con acentos aparecen como `Situaci�n` en algunas inspecciones (mojibake en herramientas, no en el Excel final).

## Glosario de dominio
- **Antigüedad de cartera**: reporte de créditos vigentes con su mora y saldos.
- **Acreditado**: cliente con crédito (`Código acreditado`, 6 díg).
- **Ciclo**: número de crédito del acreditado (1º, 2º…). Clave compuesta con código.
- **Recuperador / Promotor / Coordinación**: estructura de cobranza/territorial.
- **PAR (Portfolio at Risk)**: clasificación del riesgo por rangos de días de mora.
- **Días de mora**: días de atraso del crédito.
- **Situación crédito**: estatus del crédito. Valores vistos: `Entregado`, `Liquidado`, `Autorizado por cartera`, `Autorizado por tesorería`.
- **RECUPERADOR_000124**: recuperador específico cuyos registros se segregan (`CODIGOS_RECUPERADOR_EXCLUIR`).
- **Concepto Depósito / Saldo riesgo / Cuotas sin pagar / Combinado / Alerta**: columnas calculadas para gestión de cobranza.
- **Liquidación anticipada**: pago total adelantado; hoja con calculadora VLOOKUP.
- **Inicio ciclo**: fecha de inicio del ciclo de crédito; base de las hojas Marzo2026/Abril2026.
