# Research: Cambios al Reporte — Análisis `ReportedeAntiguedad_nuevo_31032026.xlsx`

> Fecha de análisis: 2026-04-02
> Archivo analizado: `ReportedeAntiguedad_nuevo_31032026.xlsx`
> Archivo de referencia (sistema actual): `PLANTIILA2.xlsx`

---

## 1. Resumen Ejecutivo de Cambios

| Área | Sistema actual | Sistema nuevo |
|---|---|---|
| Hojas en el output | 3 (R_Completo, X_Coordinación, X_Recuperador) | 12 hojas |
| Hoja de datos principal | `R_Completo` | `31032026` (nombre = fecha de corte) |
| Columnas de datos | 71 | 74 (3 nuevas: Cuotas sin pagar, Saldo_Riesgo_total, Combinado) |
| Categorías PAR | 7 ('PAR 0'…'PAR 6') | 8 (0, 7, 15, 30, 60, 90, Mayor_90, **Mayor_180**) |
| Manejo de fraudes | Filtrado (se excluyen del reporte) | Sin cambios — se siguen filtrando igual que ahora |
| Pivot caches | 1 (fuente: R_Completo, 69 campos) | 3 (fuentes: R_Completo + hoja de fecha con 72 y 74 campos) |
| Pivots en X_Coordinación | 2 | 6 (estructura doble: corte actual + siguiente período) |
| Pivots en X_Recuperador | 2 | 2 (sin cambios relevantes) |
| Fila TOTAL GENERAL | No existe | Sí, fila 34 en X_Coordinación suma ambas secciones |
| Hojas de origen externo | No hay | **Asignación** (oculta), **Recuperación** (oculta), **Cobranza** — las 3 son reportes externos pegados manualmente |

---

## 2. Inventario Completo de Hojas

### Hojas que Python debe generar

| Hoja | Estado | Descripción |
|---|---|---|
| `R_Completo` | **ACTUALIZADA** | Ahora 74 cols (71 base + 3 nuevas). Python la llena y todos los pivots se alimentan de ella. Mecanismo igual que hoy: tabla Excel → pivots se refrescan automáticamente. |
| `[fecha_corte]` (ej: `31032026`) | **NUEVA** | Copia idéntica de R_Completo (74 cols). Generada por Python. Nombre = fecha del corte en DDMMYYYY. |
| `[próximo_corte]` (ej: `Abril2026`) | **NUEVA** | Misma estructura (74 cols). Registros cuyo `Inicio ciclo` cae en el siguiente mes. Python la genera filtrando del mismo dataset. |
| `X_Coordinación` | **REDISEÑADA** | 6 pivots, doble sección, nueva fila TOTAL GENERAL |
| `X_Recuperador` | Sin cambios relevantes | 2 pivots, misma estructura |
| `RECUPERADOR_000124` | Sin cambios | 71 cols + 2 extras |
| `Mora` | **Reordenada** | Mismas cols pero en orden diferente (empieza en Nom. región, no Código acreditado) |
| `Cuentas con saldo vencido` | Sin cambios | 71 cols + 2 extras |
| `Liquidación anticipada` | **Cambió referencia** | VLOOKUP ahora apunta a hoja de fecha, no a R_Completo |

### Hojas que vienen de archivos externos (Python NO debe generar)

| Hoja | Origen | Evidencia |
|---|---|---|
| `Recuperación` | Reporte bancario/ERP externo | Título "REPORTE DE PAGOS / Moneda: PESO". Columnas de cuentas bancarias, referencias de pago, ciclo real. Sin fórmulas. Es un export de transacciones del sistema de administración de créditos. |
| `Cobranza` | Sistema ECONES externo | Título "Cobranza ECONES 02/Marzo - 31/Marzo". Sin fórmulas. 63 celdas mergeadas (formato de reporte manual). Es un reporte de gestión de cobranza de campo del sistema ECONES. |
| `Asignación` | **Probablemente acumulado histórico** | 717 filas = múltiples cortes. Tabla T_03022026. Tiene columna extra "Corte" al inicio. Estructura idéntica a R_Completo (71 cols) + Corte = 72 cols. Podría generarse en Python acumulando R_Completo de cada corte, pero también podría ser un archivo externo que el usuario pega manualmente. **Requiere confirmar con el usuario.** |

---

## 3. Nueva Hoja de Fecha: `[fecha_corte]` (ej: `31032026`)

Esta hoja es una **copia de R_Completo** con el mismo contenido. R_Completo sigue siendo la fuente de todos los pivots. La hoja de fecha existe como registro nombrado del corte.

### Estructura

- **Tabla Excel**: `T_02032026`
- **Header en fila 2**, datos desde fila 3
- **74 columnas**: las primeras 71 son idénticas a R_Completo + 3 nuevas (BT–BV)

> Las columnas BW en adelante (Fraude, Columna1, listas embebidas) existen en el archivo de ejemplo pero **NO se implementarán**.

### Las 3 columnas nuevas

| # Col | Letra | Nombre | Lógica | Descripción |
|---|---|---|---|---|
| 72 | BT | `Cuotas sin pagar` | `dias_ultimo_pago / dias_periodicidad` (sin redondear) | Cuotas transcurridas desde el último pago. Puede ser decimal. |
| 73 | BU | `Saldo_Riesgo_total` | `saldo_total if mora > 30 else 0` | **Nueva definición**: saldo total si mora > 30 días, sino 0. Distinto al `Saldo riesgo total` actual (suma de saldos vencidos). |
| 74 | BV | `Combinado` | `round(dias_ultimo_pago / dias_periodicidad) if mora <= 30 else saldo_riesgo_total_nuevo` | Si mora ≤ 30: cuotas sin pagar (entero). Si mora > 30: Saldo_Riesgo_total (col BU). |

### Detalle de fórmulas

**Columna BT — Cuotas sin pagar:**
```excel
=IF(Q3="Mensual",BR3/30,
 IF(Q3="Quincenal",BR3/15,
 IF(Q3="Catorcenal",BR3/14,
 IF(Q3="Semanal",BR3/7,0))))
```
Nota: sin ROUND, devuelve decimal. Q3 = Periodicidad, BR3 = Días desde el último pago.

**Columna BU — Saldo_Riesgo_total (nueva definición):**
```excel
=IF(N3>30, AY3, 0)
```
Donde AY3 = Saldo total (col 51). Esto es radicalmente diferente al cálculo actual:
- **Actual**: `Saldo riesgo total = Saldo capital + Saldo vencido + Intereses vencidos + Comisión vencida + Recargos`
- **Nuevo**: `Saldo_Riesgo_total = Saldo total si mora > 30, sino 0`

**Columna BV — Combinado:**
```excel
=IF(N3<=30,
 IF(Q3="Mensual",ROUND(BR3/30,0),
 IF(Q3="Quincenal",ROUND(BR3/15,0),
 IF(Q3="Catorcenal",ROUND(BR3/14,0),
 IF(Q3="Semanal",ROUND(BR3/7,0),0)))),
 BP3)
```
Donde BP3 = Saldo riesgo total actual (col 68). Para mora ≤ 30 devuelve cuotas (entero), para mora > 30 devuelve el monto de riesgo.

**Columna BW — Fraude:**
```excel
=XLOOKUP(BX3, CB:CB, CC:CC, 0)
```
BX3 tiene el código numérico (sin ceros). CB contiene los 27 códigos de fraude numéricos. CC contiene todo 1s.

---

## 4. Cambio en la Fórmula PAR

Este es un cambio crítico. La lógica de categorización de mora cambió.

### Sistema actual (Python, en `reportes.py`)

```python
# Categorías textuales tipo 'PAR N'
if mora == 0:    return 'PAR 0'
if mora <= 7:    return 'PAR 1'
if mora <= 15:   return 'PAR 2'
if mora <= 30:   return 'PAR 3'
if mora <= 60:   return 'PAR 4'
if mora <= 90:   return 'PAR 5'
else:            return 'PAR 6'
```
**7 categorías**. Valores tipo 'PAR 0'…'PAR 6'.

### Sistema nuevo (fórmula Excel en hoja de fecha, col O)

```excel
=IF(N3=0,"0",
 IF(AND(N3>0,N3<=7),7,
 IF(AND(N3>7,N3<=15),15,
 IF(AND(N3>15,N3<=30),30,
 IF(AND(N3>30,N3<=60),60,
 IF(AND(N3>60,N3<=90),90,
 IF(AND(N3>90,N3<=180),"Mayor_90",
 IF(N3>180,"Mayor_180",""))))))))
```
**8 categorías**. Valores: `"0"`, `7`, `15`, `30`, `60`, `90`, `"Mayor_90"`, `"Mayor_180"`.

### Diferencias concretas

| Rango mora | PAR actual | PAR nuevo |
|---|---|---|
| 0 días | 'PAR 0' | "0" (string) |
| 1–7 días | 'PAR 1' | 7 (número) |
| 8–15 días | 'PAR 2' | 15 (número) |
| 16–30 días | 'PAR 3' | 30 (número) |
| 31–60 días | 'PAR 4' | 60 (número) |
| 61–90 días | 'PAR 5' | 90 (número) |
| 91–180 días | 'PAR 6' | "Mayor_90" (string) |
| > 180 días | (no existe) | **"Mayor_180" (nuevo)** |

**Impacto**: los pivots de X_Coordinación y X_Recuperador usan la columna PAR como eje de columnas. Si Python sigue generando 'PAR 0'…'PAR 6' en lugar de 0/7/15/.../Mayor_180, los pivots mostrarán categorías incorrectas o vacías.

---

## 5. Manejo de Fraudes — Sin Cambios

Python sigue eliminando los registros con `Código acreditado` en `LISTA_FRAUDE` antes de escribir cualquier hoja. No aparecen en ningún lugar del output.

La columna `Fraude` y los datos embebidos (cols CB–CC) que aparecen en el archivo de ejemplo fueron una prueba manual y **no forman parte del diseño**.

---

## 6. Rediseño de `X_Coordinación`

### Estructura actual (PLANTIILA2.xlsx)

- 2 pivots, 1 sección de datos (filas 8–11 aprox.)
- Sin fila de TOTAL GENERAL
- Pivot 1 (A8:G16): financiero por coordinación
- Pivot 2 (J7:N11): PAR breakdown (clientes + riesgo) — 7 categorías

### Estructura nueva (31032026)

**34 filas, 54 columnas.** 6 pivots, 2 secciones + fila TOTAL GENERAL.

```
Filas 1–5:    Encabezados estáticos (PAR labels en cols AI–AV, cabeceras en A6–AF6)
Fila 6:       A=Coordinación, B=Cantidad Prestada, ... G=Saldo Riesgo Total, AF=% MORA, AI...AV=PAR clientes/riesgo
─────────────── SECCIÓN 1: CORTE ACTUAL ───────────────
Fila 7:       Cabeceras pivot medio (J7=PAR, AI7=PAR)
Fila 8:       Etiquetas PAR numéricos (J8=7, L8=15, N8=30, P8=60, R8=90, T8=Mayor_90, V8=Mayor_180)
Fila 9:       Sub-headers de los pivots
Filas 10–15:  Datos por coordinación (Atlacomulco, Maravatio, Metepec, Tenancingo, Valle de bravo, en blanco)
Fila 16:      Totales de sección 1
─────────────── SEPARADOR ───────────────
Fila 19:      Etiqueta "Abril" (label del siguiente período)
Filas 20–22:  Encabezados repetidos
Fila 23:      Cabeceras pivot medio del bloque 2
Fila 24:      Etiquetas PAR del bloque 2
Fila 25:      Sub-headers del bloque 2
Filas 26–31:  Datos por coordinación (bloque 2)
Fila 32:      Totales de sección 2
─────────────── TOTAL GENERAL ───────────────
Fila 34:      =IFERROR(SUM(X16, X32), " ")  para cada columna
```

#### Las 6 tablas dinámicas

| Pivot | Nombre | Rango | Cache (fuente) | Filas | Cols | Valores |
|---|---|---|---|---|---|---|
| 1 | TablaDinámica2 | A8:G16 | 24 → `31032026` A2:BT | [Coordinación] | [Valores] | Suma: Cantidad Prestada, Saldo Capital, Saldo Vencido, Saldo Total, Saldo Riesgo Capital, Saldo Riesgo Total |
| 2 | TablaDinámica6 | I7:AA16 | 28 → `31032026` A2:BV | [Coordinación] | [PAR] × [Valores] | Cuenta de Cuotas sin pagar + Suma de Saldo_Riesgo_total |
| 3 | TablaDinámica3 | AH7:AZ16 | 24 → `31032026` A2:BT | [Coordinación] | [PAR] × [Valores] | Cuenta de Código acreditado + Suma de Saldo riesgo total |
| 4 | TablaDinámica8 | A24:G32 | 24 → `31032026` A2:BT | [Coordinación] | [Valores] | Mismos 6 sumas que Pivot 1 (sección 2) |
| 5 | TablaDinámica7 | I23:AA32 | 28 → `31032026` A2:BV | [Coordinación] | [PAR] × [Valores] | Cuenta Cuotas sin pagar + Suma Saldo_Riesgo_total (sección 2) |
| 6 | TablaDinámica9 | AH23:AZ32 | 24 → `31032026` A2:BT | [Coordinación] | [PAR] × [Valores] | Cuenta Código acreditado + Suma Saldo riesgo total (sección 2) |

#### Columna H — fórmula nueva

```excel
H10: =J10+L10+N10
H11: =J11+L11+N11
...
```
Suma las columnas J, L y N del pivot central (TablaDinámica6), que corresponden a PAR 7, 15 y 30 (mora ≤ 30 días). Resultado = cuotas sin pagar de baja mora por coordinación.

Antes, H era `% MORA`. Ahora:
- **H** = suma de cuotas sin pagar (PAR 7+15+30) — fórmula estática
- **AF** = % MORA = `=IFERROR(D/E, 0)` — movida de H a AF

#### Columna AF — % MORA (se movió)

```excel
AF10: =IFERROR(D10/E10, 0)   → Saldo Vencido / Saldo Total
```
Igual que antes pero ahora en columna AF en lugar de H.

#### Fila 34 — TOTAL GENERAL (nueva)

```excel
C34: =IFERROR(SUM(C16, C32), " ")
D34: =IFERROR(SUM(D16, D32), " ")
...
AZ34: =IFERROR(SUM(AZ16, AZ32), " ")
```
Suma los totales de la sección 1 (fila 16) y sección 2 (fila 32) para cada columna. Abarca desde C hasta AZ (todas las columnas de datos de ambos grupos de pivots).

#### Pivot caches en la plantilla nueva (referencia al archivo de ejemplo)

En el archivo de ejemplo los caches apuntan a la hoja `31032026` porque fue editado a mano. **En la nueva plantilla que el usuario creará, todos los caches apuntarán a `R_Completo`**, que tendrá las 81 columnas completas.

| Cache (ejemplo) | Fuente en ejemplo | Fuente en plantilla nueva | Campos | Usado por |
|---|---|---|---|---|
| 23 | `R_Completo!A2:BQ` | `R_Completo!A2:BQ` | 69 cols (hasta % MORA) | X_Recuperador |
| 24 | `31032026!A2:BT` | `R_Completo!A2:BT` | 72 cols (+Cuotas sin pagar) | Pivots 1, 3, 4, 6 de X_Coordinación |
| 28 | `31032026!A2:BV` | `R_Completo!A2:BV` | 74 cols (+Saldo_Riesgo_total, +Combinado) | Pivots 2, 5 de X_Coordinación |

Todos con `refreshOnLoad="1"`.

---

## 7. `X_Recuperador` — Sin Cambios Relevantes

Sigue usando cache 23 (fuente: R_Completo) con los mismos 2 pivots y la misma estructura.

**Diferencia menor**: el pivot de PAR ahora muestra categorías `0, 7, 15, 30, 60, 90, Mayor_90, (en blanco)` en lugar de los labels de texto anteriores. Esto sugiere que también en R_Completo la columna PAR debería cambiar al nuevo formato numérico, aunque el cache sigue siendo R_Completo.

**Bug persistente**: las fórmulas de % MORA (col J) siguen desplazadas 1 fila. No fue corregido en el nuevo diseño.

---

## 8. Cambios en `Liquidación anticipada`

### Sistema actual

Python buscaba dinámicamente la hoja de fecha por nombre y construía VLOOKUPs referenciando su rango.

### Sistema nuevo (fórmulas en Excel)

Referencia explícita a `'31032026'!$A:BP`:

```excel
B3: =IFERROR(VLOOKUP(A3,'31032026'!$A:BP, 8,  FALSE),"")   → Ciclo
C3: =IFERROR(VLOOKUP(A3,'31032026'!$A:BP, 9,  FALSE),"")   → Nombre acreditado
D3: =IFERROR(VLOOKUP(A3,'31032026'!$A:BP, 26, FALSE), 0)   → Saldo interés vencido
E3: =IFERROR(VLOOKUP(A3,'31032026'!$A:BP, 27, FALSE), 0)   → Saldo comisión vencida
F3: =IFERROR(VLOOKUP(A3,'31032026'!$A:BP, 28, FALSE), 0)   → Saldo recargos
G3: =IFERROR(VLOOKUP(A3,'31032026'!$A:BP, 23, FALSE), 0)   → Saldo capital
H3: =IFERROR(VLOOKUP(A3,'31032026'!$A:BP, 52, FALSE), 0)   → Saldo adelantado
K3: =SUM(D3:G3, I3:J3) - H3                                → Cantidad a liquidar
```

Las columnas de índice son correctas y apuntan a las mismas columnas que antes. El cambio es que la referencia es a la hoja de fecha (`31032026`) en lugar de a R_Completo, ya que la hoja de fecha es la fuente de datos principal.

**Para Python**: la plantilla nueva ya tiene estas fórmulas hardcodeadas con `'31032026'`. Python debe escribir el nombre de hoja correcto al renombrar la hoja de fecha cuando genere un nuevo reporte.

---

## 9. Cambios en la Hoja `Mora`

### Reordenamiento de columnas

| Posición | Sistema actual (R_Completo order) | Sistema nuevo (Mora order) |
|---|---|---|
| A | Código acreditado | Nom. región |
| B | Nom. región | Coordinación |
| C | Coordinación | Código promotor |
| D | Código promotor | Nombre promotor |
| E | Nombre promotor | Código recuperador |
| F | Código recuperador | Nombre recuperador |
| G | Nombre recuperador | **Código acreditado** (movido a col G) |

El resto de las 71 columnas de datos sigue igual, solo cambia el orden inicial. Las columnas de seguimiento al final (Call Center + Cobranza en Campo) permanecen igual.

---

## 10. Hoja Oculta `Asignación` — Reporte Externo

**Estado**: oculta (`sheetState="hidden"`). **Python NO la genera.**

### Estructura verificada

- **Estado**: hidden
- **Dimensiones**: A2:BT717 (fila 1 vacía, headers en fila 2, datos filas 3–717 = **715 registros**)
- **Tabla Excel**: `T_03022026` (rango A2:BT717)
- **72 columnas** (A–BT)

### Listado completo de columnas (fila 2)

| Col | Letra | Nombre |
|---|---|---|
| 1 | A | Corte |
| 2 | B | Código acreditado |
| 3 | C | Nom. región |
| 4 | D | Coordinación |
| 5 | E | Código promotor |
| 6 | F | Nombre promotor |
| 7 | G | Código recuperador |
| 8 | H | Nombre recuperador |
| 9 | I | Ciclo |
| 10 | J | Nombre acreditado |
| 11 | K | Inicio ciclo |
| 12 | L | Fin ciclo |
| 13 | M | Cantidad entregada |
| 14 | N | Cantidad Prestada |
| 15 | O | Días de mora |
| 16 | P | PAR |
| 17 | Q | Plazo del crédito |
| 18 | R | Periodicidad |
| 19 | S | Parcialidad + Parcialidad comisión |
| 20 | T | Comisión a pagar |
| 21 | U | Interés moratorio saldo vencido total |
| 22 | V | Pagos vencidos |
| 23 | W | Periodos vencidos |
| 24 | X | Saldo capital |
| 25 | Y | Saldo vencido |
| 26 | Z | Saldo capital vencido |
| 27 | AA | Saldo interés vencido |
| 28 | AB | Saldo comisión vencida |
| 29 | AC | Saldo recargos |
| 30 | AD | $ último pago |
| 31 | AE | Último pago |
| 32 | AF | Situación crédito |
| 33 | AG | Medio comunic. 1 |
| 34 | AH | Medio comunic. 2 |
| 35 | AI | Medio comunic. 3 |
| 36 | AJ | Actividad económica PLD |
| 37 | AK | Código actividad económica PLD |
| 38 | AL | Nombre conyuge |
| 39 | AM | Teléfono conyuge |
| 40 | AN | Nombre Referencia1 |
| 41 | AO | Teléfono Referencia1 |
| 42 | AP | Nombre Referencia2 |
| 43 | AQ | Teléfono Referencia2 |
| 44 | AR | Nombre Referencia3 |
| 45 | AS | Teléfono Referencia3 |
| 46 | AT | Tipo Garantía 1 |
| 47 | AU | Descripción garantía 1 |
| 48 | AV | Garantía 1 |
| 49 | AW | Tipo Garantía 2 |
| 50 | AX | Descripción garantía 2 |
| 51 | AY | Garantía 2 |
| 52 | AZ | Saldo total |
| 53 | BA | Saldo adelantado |
| 54 | BB | Cód. producto crédito |
| 55 | BC | Nom. producto crédito |
| 56 | BD | Calle |
| 57 | BE | Colonia |
| 58 | BF | Entidad federativa |
| 59 | BG | Municipio |
| 60 | BH | Geolocalización domicilio |
| 61 | BI | Link de Geolocalización ← **única columna con fórmulas** |
| 62 | BJ | Castigado cartera |
| 63 | BK | Nom. personal castiga cartera |
| 64 | BL | Frecuencia |
| 65 | BM | Criticidad |
| 66 | BN | Forma de entrega |
| 67 | BO | Concepto Depósito |
| 68 | BP | Saldo riesgo capital |
| 69 | BQ | Saldo riesgo total |
| 70 | BR | % MORA |
| 71 | BS | Columna1 |
| 72 | BT | Columna2 |

### Fórmulas

Solo la columna BI (`Link de Geolocalización`) contiene fórmulas HYPERLINK:
```excel
=HYPERLINK("https://maps.google.com","Ver en mapa")          ← la mayoría (sin coords reales)
=HYPERLINK("https://www.google.com/maps/search/?api=1&query=19.167805555555557,-100.21975","Ver en mapa")  ← cuando sí hay coords
```
El resto de la hoja es datos estáticos (no hay otras fórmulas).

### Origen

Es un **Reporte de Antigüedad de Cartera** exportado del sistema de administración de créditos — el mismo tipo de reporte que sirve como input a nuestro sistema Python, pero con más columnas al final (Frecuencia, Criticidad, Forma de entrega, Concepto Depósito, Saldo riesgo capital, Saldo riesgo total, % MORA, Columna1, Columna2). Estas columnas extra NO existen en el R_Completo que Python genera actualmente.

**Conclusión**: es un export directo del sistema, pegado manualmente en el workbook. **Python no debe generarla.**

### Uso en Cobranza

La hoja Cobranza tiene esta fórmula que la referencia:
```excel
=XLOOKUP(F8,'31032026'!$A:$A,'31032026'!$G:$G, XLOOKUP(F8, Asignación!$B:$B, Asignación!$H:$H,))
```
- Busca el **Nombre recuperador** (col H) por **Código acreditado** (col B) como fallback si no lo encuentra en la hoja de fecha.

---

## 11. Hoja Oculta `Recuperación` — Reporte Externo de Pagos

**Estado**: oculta (`sheetState="hidden"`). **Python NO la genera.**

### Estructura verificada

- **Dimensiones**: A1:CW155
- **Fila 1**: título "REPORTE DE PAGOS" en col E
- **Fila 2**: "Moneda: PESO" en col E
- **Fila 3**: headers de columnas
- **Filas 4–155**: 152 registros de pagos
- **Sin fórmulas, sin tablas Excel**
- **101 columnas declaradas** (A–CW), aunque solo A–BX (76 cols) tienen encabezado con datos

### Listado completo de columnas (fila 3)

| Col | Letra | Nombre |
|---|---|---|
| 1 | A | Fecha corte ← usada en SUMIFS de Cobranza |
| 2 | B | Fecha del depósito |
| 3 | C | Fecha de pago |
| 4 | D | Código ← usada en SUMIFS de Cobranza |
| 5 | E | Acreditados |
| 6 | F | Ciclo |
| 7 | G | Ciclo real |
| 8 | H | Periodo |
| 9 | I | Referencia |
| 10 | J | Modo |
| 11 | K | Estatus |
| 12 | L | Forma de pago |
| 13 | M | Forma de distribución de interés |
| 14 | N | Código cta ban. |
| 15 | O | No. cuenta bancaria |
| 16 | P | Cuenta bancaria |
| 17 | Q | Núm. cheque |
| 18 | R | Cantidad ← usada en SUMIFS de Cobranza (monto del pago) |
| 19 | S | Capital |
| 20 | T | Interés |
| 21 | U | Recargos |
| 22 | V | Comisión |
| 23 | W | Cargo Seguro de vida |
| 24 | X | Abono Seguro de vida |
| 25 | Y | Cargo Seguro Auto |
| 26 | Z | Abono Seguro Auto |
| 27 | AA | Cargo GPS |
| 28 | AB | Abono GPS |
| 29 | AC | Cargo Multas por cheque rebotado |
| 30 | AD | Abono Multas por cheque rebotado |
| 31 | AE | Dep. en garantía |
| 32 | AF | Ahorro |
| 33 | AG | Cód. región |
| 34 | AH | Región |
| 35 | AI | Cód. coordinación |
| 36 | AJ | Coordinación |
| 37 | AK | Cód. tipo de empresa |
| 38 | AL | Tipo de Empresa |
| 39 | AM | Cód. empresa |
| 40 | AN | Empresa |
| 41 | AO | Cód. empresa cobro |
| 42 | AP | Empresa cobro |
| 43 | AQ | Cód. puntos de venta |
| 44 | AR | Puntos de venta |
| 45 | AS | Cód. sector |
| 46 | AT | Sector |
| 47 | AU | Cód. giro |
| 48 | AV | Giro |
| 49 | AW | Tipo crédito |
| 50 | AX | Tipo préstamo |
| 51 | AY | Municipio |
| 52 | AZ | Cód. org. fond. |
| 53 | BA | Organización fondeadora |
| 54 | BB | Cód. línea de crédito |
| 55 | BC | Línea de crédito |
| 56 | BD | Conciliado |
| 57 | BE | Tipo |
| 58 | BF | No. factura |
| 59 | BG | Cód. Recuperador |
| 60 | BH | Recuperador |
| 61 | BI | Fecha de registro |
| 62 | BJ | Cod. tipo producto de crédito |
| 63 | BK | Nombre del tipo producto de crédito |
| 64 | BL | Cod. producto de crédito |
| 65 | BM | Nombre del producto de crédito |
| 66 | BN | No. caja |
| 67 | BO | CIE |
| 68 | BP | Folio |
| 69 | BQ | RFC |
| 70 | BR | Folio préstamo |
| 71 | BS | Número traspaso fondo |
| 72 | BT | Número de empleado |
| 73 | BU | Fecha de castigo |
| 74 | BV | Plazo |
| 75 | BW | Periodicidad |
| 76 | BX | Fecha movimiento bancario |

### Origen

**"REPORTE DE PAGOS"** del sistema administrador de créditos — export de transacciones de pago por acreditado. Es el mismo sistema que genera el Reporte de Antigüedad de Cartera. El usuario lo pega manualmente.

### Uso en Cobranza

La hoja Cobranza lo referencia en esta fórmula SUMIFS:
```excel
=SUMIFS(Recuperación!$R:$R, Recuperación!$D:$D, $F8, Recuperación!$A:$A, $Q$7)
```
- Suma **Cantidad** (col R = monto del pago) filtrada por **Código** (col D) y **Fecha corte** (col A).
- Equivale a: "¿Cuánto pagó este acreditado en este corte?"

---

## 12. Nueva Hoja `Cobranza` — Origen Externo Confirmado

**No debe generarse en Python.**

- **212 filas × 22 columnas**
- Título: "Cobranza ECONES 02/Marzo - 31/Marzo"
- Sin fórmulas, sin tablas Excel
- 63 celdas mergeadas (formato de reporte de gestión)
- **Conclusión**: reporte del sistema ECONES (sistema de gestión de cobranza en campo). Se pega manualmente en el workbook.

---

## 13. Nueva Hoja del Siguiente Período (ej: `Abril2026`)

- **Misma estructura que R_Completo y la hoja de fecha (81 cols)**
- Contiene los registros del dataset cuyo campo `Inicio ciclo` cae en el siguiente mes al corte
- En el ejemplo: 6 registros con `Inicio ciclo = 02/04/2026`
- Python la genera filtrando `df[df['Inicio ciclo'].dt.month == mes_siguiente]`
- El nombre de la hoja = nombre del mes + año del siguiente período (ej: `Abril2026`)

Lleva las mismas 74 cols que R_Completo (71 base + Cuotas sin pagar + Saldo_Riesgo_total + Combinado). Sin columnas de fraude.

---

## 14. Mapa Completo de lo que Python Debe Implementar

### Columnas calculadas nuevas (cols BT–BV, aplicadas a R_Completo y hojas de fecha)

| Columna | Nombre | Lógica Python equivalente |
|---|---|---|
| BT | Cuotas sin pagar | `dias_ultimo_pago / dias_periodicidad` (sin redondear, decimal) |
| BU | Saldo_Riesgo_total | `saldo_total if mora > 30 else 0` |
| BV | Combinado | `round(dias_ultimo_pago / dias_periodicidad) if mora <= 30 else saldo_riesgo_total_bu` |

> Columnas BW en adelante (Fraude, Columna1, listas embebidas) **no se implementan**.

### Manejo de fraudes

Sin cambios. El sistema sigue filtrando los registros de `LISTA_FRAUDE` antes de escribir cualquier hoja. La columna `Fraude` que aparece en el archivo de ejemplo fue una prueba manual, no forma parte del diseño.

### Cambios en PAR

```python
def asignar_par_nuevo(dias_mora):
    if dias_mora == 0:        return "0"
    if dias_mora <= 7:        return 7
    if dias_mora <= 15:       return 15
    if dias_mora <= 30:       return 30
    if dias_mora <= 60:       return 60
    if dias_mora <= 90:       return 90
    if dias_mora <= 180:      return "Mayor_90"
    return "Mayor_180"        # NUEVA CATEGORÍA
```

### Nueva estructura del output Excel

```
1. R_Completo          ← ACTUALIZADO: 74 cols, fuente permanente de todos los pivots
2. [fecha_corte]       ← NUEVO: copia idéntica de R_Completo (74 cols), nombre = DDMMYYYY
3. [próximo_período]   ← NUEVO: misma estructura (74 cols), filtrado por Inicio ciclo del mes siguiente
4. X_Coordinación      ← REDISEÑADO: 6 pivots, doble sección, TOTAL GENERAL (fuente: R_Completo)
5. X_Recuperador       ← sin cambios relevantes (fuente: R_Completo, igual que hoy)
6. RECUPERADOR_000124  ← sin cambios
7. Mora                ← mismo contenido, columnas reordenadas
8. Cuentas con saldo vencido ← sin cambios
9. Liquidación anticipada   ← fórmulas apuntan a R_Completo (Python actualiza referencia)
10. Recuperación       ← NO generar (externo, pegado manualmente)
11. Cobranza           ← NO generar (externo, pegado manualmente)
12. Asignación         ← NO generar (externo, pegado manualmente)
```

---

## 15. Eliminación de Hojas por Coordinación

En el sistema actual, Python genera **una hoja por cada coordinación única** encontrada en los datos (ej: "Atlacomulco", "Maravatio", "Metepec", "Tenancingo", "Valle de bravo"). Cada hoja contiene todos los registros filtrados de esa coordinación con formato completo.

En el nuevo reporte **estas hojas no existen**. El inventario de 12 hojas no incluye ninguna hoja con nombre de coordinación.

**Implicación directa para Python**: eliminar completamente la lógica que distribuye registros por coordinación y crea una hoja por cada una. Todo el desglose por coordinación ahora vive en los pivots de `X_Coordinación`.

---

## 16. Estado de Decisiones y Puntos Pendientes

### Confirmado por el usuario

| # | Decisión | Estado |
|---|---|---|
| 1 | **Asignación** es externa — el usuario la pega manualmente | ✅ Confirmado |
| 2 | **Recuperación** es externa — reporte de pagos del sistema | ✅ Confirmado (estructura verificada) |
| 3 | **Cobranza** es externa — reporte ECONES, se pega manualmente | ✅ Confirmado |
| 4 | **Fraudes** se siguen filtrando igual que ahora. La columna Fraude del ejemplo fue prueba manual. | ✅ Confirmado |
| 5 | **Abril2026** Python la genera filtrando por Inicio ciclo del siguiente mes | ✅ Confirmado (todos los registros tienen Inicio ciclo 02/04/2026) |
| 6 | **R_Completo se actualiza a 81 cols** (mismas que hoja de fecha). Todos los pivots la usan como fuente. El usuario creará nueva plantilla con los rangos correctos. | ✅ Confirmado |
| 7 | **Nombre hoja de fecha** se genera automáticamente igual que ahora (DDMMYYYY) | ✅ Confirmado |
| 8 | **Liquidación anticipada** — Python ya escribe fórmulas dinámicas con `nombre_hoja_informe`; comportamiento actual se mantiene | ✅ Verificado en código |

### Pendiente de confirmar

1. **`Mora` reordenada**: ¿el nuevo orden de columnas (empieza por Nom. región en lugar de Código acreditado) es intencional?
3. **Sección 2 de X_Coordinación ("Abril")**: los pivots no tienen filtro aplicado en el ejemplo. ¿Se implementará con opción 2 (page filter por Inicio ciclo) u otra variante? Marcado como probable opción 2 pero pendiente de confirmar.
