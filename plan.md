# Plan de Implementación — Nuevo Formato de Reportes

> Estrategia: cambios incrementales, uno a la vez. Cada iteración se valida antes de continuar.
> Si una validación falla → se revierte el cambio y se ajusta antes de avanzar.

---

## Estado actual del output

```
R_Completo                  ← 71 cols
X_Coordinación              ← 2 pivots (fuente: R_Completo)
X_Recuperador               ← 2 pivots (fuente: R_Completo)
RECUPERADOR_000124          ← 71 cols + extras
Mora                        ← orden actual (empieza con Código acreditado)
Cuentas con saldo vencido   ← sin cambios
Liquidación anticipada      ← fórmulas apuntan a R_Completo
Atlacomulco                 ← ELIMINAR
Maravatio                   ← ELIMINAR
Metepec                     ← ELIMINAR
Tenancingo                  ← ELIMINAR
Valle de bravo              ← ELIMINAR
```

## Estado objetivo del output

```
R_Completo                  ← 74 cols (71 + 3 nuevas)
[DDMMYYYY]  (ej: 31032026)  ← copia idéntica de R_Completo
[MesAño]    (ej: Abril2026) ← mismas 74 cols, filtrado por Inicio ciclo siguiente mes
X_Coordinación              ← 6 pivots (lo modifica el usuario en la plantilla)
X_Recuperador               ← 2 pivots (sin cambios estructurales)
RECUPERADOR_000124          ← sin cambios
Mora                        ← reordenada (empieza con Nom. región)
Cuentas con saldo vencido   ← sin cambios
Liquidación anticipada      ← sin cambios (Python ya escribe fórmulas dinámicas)
```

---

## Iteración 1 — Eliminar hojas por coordinación

**Objetivo**: quitar las hojas `Atlacomulco`, `Maravatio`, `Metepec`, `Tenancingo`, `Valle de bravo` del output.

**Archivo a modificar**: `app/reportes.py`

**Cambio técnico**:
- Buscar la sección que itera por coordinaciones y crea una hoja por cada una
- Comentar o eliminar ese bloque completo
- No tocar nada más

**Cómo validar**:
1. Subir cualquier reporte de antigüedad al sistema
2. Descargar el Excel generado
3. Confirmar que el archivo ya NO contiene hojas con nombres de coordinación
4. Confirmar que R_Completo, X_Coordinación, X_Recuperador, Mora, etc. siguen presentes y correctas

**Cómo revertir**: descomentar el bloque eliminado

---

## Iteración 2 — Cambiar fórmula PAR

**Objetivo**: actualizar la clasificación de mora de 7 categorías ('PAR 0'…'PAR 6') a 8 categorías numéricas.

**Archivo a modificar**: `app/reportes.py`

**Cambio técnico**:
Reemplazar la función que calcula PAR con la nueva lógica:

```python
# ANTES
def calcular_par(mora):
    if mora == 0:   return 'PAR 0'
    if mora <= 7:   return 'PAR 1'
    if mora <= 15:  return 'PAR 2'
    if mora <= 30:  return 'PAR 3'
    if mora <= 60:  return 'PAR 4'
    if mora <= 90:  return 'PAR 5'
    return 'PAR 6'

# DESPUÉS
def calcular_par(mora):
    if mora == 0:        return "0"
    if mora <= 7:        return 7
    if mora <= 15:       return 15
    if mora <= 30:       return 30
    if mora <= 60:       return 60
    if mora <= 90:       return 90
    if mora <= 180:      return "Mayor_90"
    return "Mayor_180"
```

**Cómo validar**:
1. Generar reporte con el sistema
2. Abrir R_Completo, buscar la columna `PAR`
3. Revisar ~5 registros con distintos días de mora y confirmar que el valor PAR corresponde a la nueva escala
4. Confirmar que los pivots de X_Recuperador reflejan las nuevas categorías al refrescar

**Cómo revertir**: restaurar la función anterior

---

## Iteración 3 — Agregar 3 columnas nuevas a R_Completo

**Objetivo**: agregar al final de R_Completo las columnas `Cuotas sin pagar` (BT), `Saldo_Riesgo_total` (BU) y `Combinado` (BV).

**Archivo a modificar**: `app/reportes.py` y posiblemente `config.py`

**Lógica de cada columna**:

```python
# Requiere: 'Días desde el último pago', 'Periodicidad', 'Días de mora', 'Saldo total', 'Saldo riesgo total'
# PERIODICIDAD_A_DIAS ya existe en config.py: {'semanal':7, 'catorcenal':14, 'quincenal':15, 'mensual':30, ...}

def calcular_cuotas_sin_pagar(dias_ultimo_pago, periodicidad):
    dias = PERIODICIDAD_A_DIAS.get(str(periodicidad).lower(), 0)
    if dias == 0:
        return 0
    return dias_ultimo_pago / dias   # decimal, sin redondear

def calcular_saldo_riesgo_total_nuevo(mora, saldo_total):
    return saldo_total if mora > 30 else 0

def calcular_combinado(mora, dias_ultimo_pago, periodicidad, saldo_riesgo_total_nuevo):
    dias = PERIODICIDAD_A_DIAS.get(str(periodicidad).lower(), 0)
    if mora <= 30:
        return round(dias_ultimo_pago / dias) if dias > 0 else 0
    else:
        return saldo_riesgo_total_nuevo
```

**Nota importante**: `Saldo_Riesgo_total` en la col BU es una **nueva definición** distinta a `Saldo riesgo total` (col ~BN) que ya existe. La del BU = `saldo_total if mora > 30 else 0`. La existente = suma de saldos vencidos. Ambas coexisten.

**Cómo validar**:
1. Generar reporte
2. En R_Completo verificar que existen 3 columnas nuevas al final: `Cuotas sin pagar`, `Saldo_Riesgo_total`, `Combinado`
3. Tomar 3-5 registros y calcular manualmente los valores para confirmar la lógica
4. Verificar que el total de columnas en R_Completo ahora es 74

**Cómo revertir**: eliminar las 3 columnas del código de escritura de R_Completo

---

## Iteración 4 — Generar hoja de fecha (copia de R_Completo)

**Objetivo**: Python genera una segunda hoja con nombre = fecha del corte en formato DDMMYYYY, con el mismo contenido que R_Completo.

**Archivo a modificar**: `app/reportes.py`

**Cambio técnico**:
- Después de escribir R_Completo, copiar el mismo DataFrame a una nueva hoja
- El nombre de la hoja se obtiene del campo de fecha del reporte (mismo mecanismo que ya usa `nombre_hoja_informe`)
- La tabla de la hoja de fecha tendrá un nombre distinto (ej: `T_DDMMYYYY`)
- Mismas 74 columnas, mismo formato

**Cómo validar**:
1. Generar reporte con archivo del corte 31/03/2026
2. Confirmar que el Excel tiene una hoja llamada `31032026`
3. Confirmar que tiene exactamente el mismo número de filas y columnas que R_Completo
4. Comparar 5 filas al azar entre R_Completo y `31032026` — deben ser idénticas

**Cómo revertir**: eliminar el bloque que escribe la hoja de fecha

---

## Iteración 5 — Generar hoja del siguiente período

**Objetivo**: Python genera una hoja con nombre del siguiente mes (ej: `Abril2026`) que contiene los registros cuyo `Inicio ciclo` cae en el mes siguiente al corte.

**Archivo a modificar**: `app/reportes.py`

**Cambio técnico**:
```python
import calendar

# Calcular el mes siguiente
mes_siguiente = (fecha_corte.month % 12) + 1
anio_siguiente = fecha_corte.year + (1 if mes_siguiente == 1 else 0)

# Filtrar registros
df_siguiente = df[df['Inicio ciclo'].dt.month == mes_siguiente]

# Nombre de la hoja: nombre del mes en español + año
MESES_ES = {1:'Enero',2:'Febrero',3:'Marzo',4:'Abril',5:'Mayo',6:'Junio',
            7:'Julio',8:'Agosto',9:'Septiembre',10:'Octubre',11:'Noviembre',12:'Diciembre'}
nombre_hoja_siguiente = f"{MESES_ES[mes_siguiente]}{anio_siguiente}"

# Escribir con mismas 74 cols y mismo formato que R_Completo
```

**Cómo validar**:
1. Generar reporte con archivo que tenga al menos un registro con `Inicio ciclo` en el mes siguiente
2. Confirmar que aparece hoja `Abril2026` (o el mes que corresponda)
3. Confirmar que solo contiene registros cuyo Inicio ciclo es en ese mes
4. Si no hay registros del mes siguiente, la hoja debe crearse vacía (solo headers) o no crearse — **definir comportamiento**

**Cómo revertir**: eliminar el bloque que escribe la hoja del siguiente período

---

## Iteración 6 — Reordenar columnas en hoja Mora

**Objetivo**: cambiar el orden inicial de columnas en la hoja `Mora` para que empiece con `Nom. región` en lugar de `Código acreditado`.

**Nuevo orden de primeras 7 columnas**:
```
A: Nom. región
B: Coordinación
C: Código promotor
D: Nombre promotor
E: Código recuperador
F: Nombre recuperador
G: Código acreditado   ← movido de A a G
```
El resto de columnas (8 en adelante) permanece igual.

**Archivo a modificar**: `app/reportes.py` — sección que escribe la hoja Mora

**Cómo validar**:
1. Generar reporte
2. Abrir hoja `Mora`
3. Confirmar que col A = `Nom. región` y col G = `Código acreditado`
4. Confirmar que el número total de filas y columnas es correcto

**Cómo revertir**: restaurar el orden original de columnas

---

## Iteración 7 — Ajustar Liquidación anticipada

**Objetivo**: verificar y ajustar que las fórmulas VLOOKUP en `Liquidación anticipada` apunten correctamente a `R_Completo` con el nuevo rango de 74 columnas.

**Contexto**: Python ya escribe estas fórmulas dinámicamente. Solo verificar que los índices de columna siguen siendo correctos con las 3 nuevas columnas al final (no afectan los índices existentes ya que se agregan al final).

**Archivo a modificar**: `app/reportes.py` — sección de Liquidación anticipada

**Qué revisar**:
- Los VLOOKUP usan índices como 8, 9, 23, 26, 27, 28, 52 — todos menores a 71, así que las 3 nuevas columnas en posición 72-74 no los afectan
- El rango del VLOOKUP debe actualizarse de `$A:BP` a `$A:BV` para consistencia

**Cómo validar**:
1. En la hoja `Liquidación anticipada` ingresar manualmente un Código acreditado en col A
2. Confirmar que las celdas B-K se rellenan con los datos correctos del acreditado
3. Verificar que `Cantidad a liquidar` (col K) calcula correctamente

**Cómo revertir**: ajustar los rangos de VLOOKUP de vuelta

---

## Iteración 8 — Crear nueva plantilla con pivots actualizados

**Objetivo**: tener una plantilla Excel con la estructura del nuevo diseño lista para que Python la use como base.

---

### Paso 8A — Intento automatizado (Python limpia el archivo nuevo)

**Lo que hace el script**:
1. Copia `ReportedeAntiguedad_nuevo_31032026.xlsx` → `PLANTILLA_NUEVA.xlsx`
2. Vacía los datos de las hojas de datos (R_Completo, 31032026, Abril2026) — deja solo headers y definición de tabla
3. Elimina las hojas externas: `Asignación`, `Recuperación`, `Cobranza`
4. Redirige los pivot caches de `31032026` → `R_Completo` directamente en el XML
5. Actualiza las fórmulas de `Liquidación anticipada` que referencian `'31032026'` → `R_Completo`

**Riesgo**: openpyxl no soporta pivots nativamente — se manipulan directo en XML. Puede que algún pivot quede corrupto.

**El archivo original `ReportedeAntiguedad_nuevo_31032026.xlsx` no se toca** — el script trabaja sobre una copia.

**Cómo validar**:
1. Abrir `PLANTILLA_NUEVA.xlsx` en Excel
2. Confirmar que las hojas externas (Asignación, Recuperación, Cobranza) ya no están
3. Confirmar que R_Completo existe pero está vacía (solo headers)
4. Ir a X_Coordinación → clic derecho en cada pivot → "Actualizar" → deben cargar sin error
5. Ir a X_Recuperador → mismo proceso
6. Confirmar que los pivot caches ahora apuntan a R_Completo (al actualizar no piden buscar fuente)

**Si el paso 8A funciona** → continuar con iteración 9.

---

### Paso 8B — Ajuste manual por el usuario (si 8A falla parcialmente)

Si los pivots quedaron corruptos o las referencias incorrectas, el usuario corrige manualmente en Excel:

1. **Ampliar el cache de X_Coordinación**:
   - Cambiar fuente de `R_Completo!A2:BQ` → `R_Completo!A2:BV`
   - Esto expone los 3 campos nuevos: `Cuotas sin pagar`, `Saldo_Riesgo_total`, `Combinado`

2. **Agregar 4 pivots nuevos** (duplicar los existentes para la Sección 2):
   - Copiar TablaDinámica2 → posición A24:G32
   - Copiar TablaDinámica3 → posición AH7:AZ16 y AH23:AZ32
   - Crear TablaDinámica6 en I7:AA16: filas=[Coordinación], cols=[PAR × VALUES], valores=[Cuenta Cuotas sin pagar + Suma Saldo_Riesgo_total]
   - Crear TablaDinámica7 en I23:AA32: igual que TablaDinámica6

3. **Agregar fila TOTAL GENERAL** en fila 34:
   - Fórmula: `=IFERROR(SUM(X16, X32), " ")` para cada columna de datos

4. **Agregar etiquetas estáticas** de PAR en filas 7-8 y 23-24 (valores: 7, 15, 30, 60, 90, Mayor_90, Mayor_180)

**Notas**:
- Los pivots de X_Recuperador no necesitan cambios estructurales
- Si la Sección 2 debe mostrar datos del siguiente período con filtro (opción 2), agregar page filter `Inicio ciclo = mes siguiente` a los pivots 4-6 — **pendiente confirmar**

---

## Iteración 9 — Validación integral y ajustes finales

**Objetivo**: prueba end-to-end con un reporte de antigüedad real. Verificar que todo el output es correcto.

**Checklist de validación**:
- [ ] R_Completo: 74 columnas, datos correctos, sin fraudes
- [ ] [DDMMYYYY]: idéntico a R_Completo
- [ ] [MesAño]: registros filtrados por Inicio ciclo del siguiente mes
- [ ] X_Coordinación: 6 pivots con datos, fila TOTAL GENERAL correcta
- [ ] X_Recuperador: 2 pivots, datos correctos
- [ ] RECUPERADOR_000124: sin cambios
- [ ] Mora: columnas en nuevo orden
- [ ] Cuentas con saldo vencido: sin cambios
- [ ] Liquidación anticipada: fórmulas funcionan al ingresar un código
- [ ] No aparecen hojas por coordinación
- [ ] Columna PAR: valores en nuevo formato (0, 7, 15, ..., Mayor_180)
- [ ] Columnas nuevas (BT-BV): calculadas correctamente

---

## Referencia rápida de archivos

| Archivo | Rol |
|---|---|
| `app/reportes.py` | Lógica principal de generación de Excel |
| `config.py` | Constantes: LISTA_FRAUDE, PERIODICIDAD_A_DIAS, COLUMN_MAPPING, etc. |
| `PLANTIILA2.xlsx` | Plantilla actual (se reemplazará con nueva plantilla del usuario) |
| `research_cambios.md` | Documentación detallada de todos los cambios |
| `research.md` | Documentación del sistema actual |
