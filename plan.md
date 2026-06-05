# CREDIFLEXI / automatizador-crediflexi — Plan / Hoja de ruta

> Documento vivo. Última actualización: 2026-06-04 (commit `fe8d8e0`).
> Contexto completo en `research.md`. Detalle histórico de iteraciones del nuevo formato en `research_cambios.md`.
> Regla: derivar el backlog del estado real del código (lo que falta para el siguiente hito), ordenado por dependencia.

## Objetivo actual
Cerrar la deuda recurrente que rompe el reporte cada mes y endurecer el flujo individual (ya maduro) antes de retomar el flujo grupal (stub). En concreto: que el reporte mensual NO requiera editar código cada corte, y dejar configurable el criterio de la hoja "Autorizado por Cartera" que quedó hardcodeado.

## Siguiente paso inmediato
**Des-hardcodear el corte mensual** (`Marzo2026`, `Abril2026`, fecha `2026-04-01`) en `app/reportes.py`.
- Hoy: nombres de hoja y fecha de corte están fijos en el código (`:1705`, `:1777`) y atados a los pivot caches de la plantilla (nombres fijos `Marzo2026`/`Abril2026`).
- Acción mínima: derivar nombre de hoja y fecha de corte de la fecha del reporte (mes actual = histórico, mes siguiente = nuevo), centralizando en una sola variable/constante.
- Restricción dura: los pivot caches de `PLANTIILA2.xlsx` apuntan a esos nombres por XML → si se renombra la hoja sin tocar la plantilla, los pivots se rompen. Decidir antes: (a) mantener nombres fijos y solo parametrizar la FECHA de corte, o (b) renombrar también en el XML de la plantilla. Confirmar con el dueño cuál.
- Criterio de aceptación: subir un input de un mes distinto a abril y obtener las dos hojas con el nombre/mes correcto sin editar código, con los pivots vivos.

## Backlog ordenado (por dependencia)
- [ ] **Parametrizar fecha de corte mensual** (siguiente paso). Aceptación: ver arriba. Bloquea cualquier corrida de un mes ≠ abril 2026.
- [ ] **Decidir y aplicar criterio configurable de "Autorizado por Cartera"**. Hoy hardcodeado a `Situación crédito ≠ Entregado` (`:1341`). Opciones abiertas (pendiente decisión del dueño): solo `Autorizado por cartera`; `Autorizado por cartera` + `Autorizado por tesorería`; lista configurable en `config.py`. Aceptación: el set de estatus que va a la hoja se lee de `config.py`, no del código; un test manual confirma el conteo esperado (hoy: 128 de 342 en el input de ejemplo).
- [ ] **Crear `requirements.txt`** con versiones fijadas (flask, flask-login, flask-sqlalchemy, flask-wtf, pandas, openpyxl, werkzeug). Aceptación: `pip install -r requirements.txt` en un venv limpio levanta la app. El README ya lo menciona pero el archivo no existe.
- [ ] **Mover `SECRET_KEY` y contraseñas semilla a variables de entorno** (`config.py`, `app/auth.py`). Aceptación: no quedan secretos en el repo; la app lee de env con fallback solo para dev. Depende de tener `requirements.txt`/onboarding claro para documentar las env vars.
- [ ] **Sincronizar README con el estado real**: puerto 5001 (no 5000), retirar mención de hojas por coordinación y del reporte de colaboradores (ya eliminados), documentar `parches.json` y la hoja "Autorizado por Cartera". Aceptación: README sin features inexistentes.
- [ ] **Endurecer recorte de columnas de R_Completo**: `df_r_completo.iloc[:, :74]` por posición fija es frágil (`research.md` deuda). Aceptación: seleccionar por nombre de columna o validar el conteo y fallar con mensaje claro si el input cambia de forma.
- [ ] **Tests mínimos del pipeline individual**: al menos un test que corra `procesar_reporte_antiguedad()` sobre `test_abril2026_input.xlsx` y verifique nº de hojas, nº de columnas de R_Completo y conteos de los subsets (fraude/RECUPERADOR_000124/Autorizado por Cartera). Aceptación: `pytest` pasa en verde. Depende de `requirements.txt`.

## Bloqueado / esperando
- **Flujo grupal** (`procesar_antiguedad_grupal`, `:3141`): STUB que genera un Excel dummy de 1 hoja. Detecta 5 tipos de archivo (`detectar_tipo_archivo`, `:3111`) pero no consolida nada. **BLOQUEADO**: falta la especificación del dueño — qué columnas/hojas debe producir, qué cruces hacer entre los 5 archivos y ejemplos de input. No avanzar sin spec.
- **Renombrado de hojas mensuales en la plantilla**: esperando decisión (a) vs (b) del "siguiente paso inmediato" antes de tocar el XML de pivots de `PLANTIILA2.xlsx`.
- **Criterio de "Autorizado por Cartera"**: esperando decisión del dueño entre las opciones listadas en el backlog.

## Hecho recientemente
- **`fe8d8e0`** — Hoja "Autorizado por Cartera": separa registros con `Situación crédito ≠ Entregado` a su propia hoja, mismo pipeline que RECUPERADOR_000124.
- **`dcc07bb`** — Sistema `parches.json`: correcciones manuales por `codigo`+`ciclo` sin tocar código (1er parche: reasignar `Coordinación`=`Oficina Central` a 18 acreditados).
- **`1b63eed`** — Abril2026 captura todos los registros con `Inicio ciclo` ≥ 2026-04-01.
- **`104e154`** — Hoja histórica Marzo2026 (registros antes de abril 2026).
- **`021445...` / `021449e`** — Columna `Suma` (col 75): 1 si días de mora ∈ [1,30].
- **`f3bace7`** — Recorte de R_Completo a 74 columnas (quita columnas basura del origen).
- **`research.md`** reescrito como doc vivo con el mapeo completo INPUT → TRANSFORMACIONES → OUTPUT.
