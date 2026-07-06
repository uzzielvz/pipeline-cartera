# CREDIFLEXI / automatizador-crediflexi — Plan / Hoja de ruta

> Documento vivo. Última actualización: 2026-07-06 (commit `acbf9e4`).
> Contexto completo en `research.md`. Detalle histórico de iteraciones del nuevo formato en `research_cambios.md`.
> Regla: derivar el backlog del estado real del código (lo que falta para el siguiente hito), ordenado por dependencia.

## Objetivo actual
Implementar 5 modificaciones solicitadas por Uzziel (2026-07-06):
1. Crear hoja "SinCastigoMarzo" (Marzo sin registros PAR Mayor_180)
2. Agregar columna "Fecha próximo pago" desde Rep_Cobranza (Yunius)
3. Mover créditos Oficina Central de Abril → Marzo
4. Parche cliente 001388: Coordinación → Metepec (preventivo)
5. Parche cliente 001053: limpiar datos personales (PII)

## Siguiente paso inmediato
**Crear hoja SinCastigoMarzo** en `app/reportes.py`:
- Clonar lógica de Marzo2026 (filtro `Inicio ciclo` < fecha corte)
- Filtro adicional: excluir registros con `PAR` = `Mayor_180`
- Insertar hoja después de Marzo2026 en ORDEN_HOJAS

## Backlog ordenado (por dependencia)
### Sprint actual (julio 2026)
- [x] **Parche cliente 001388** (Coordinación → Metepec). Commit `acbf9e4`.
- [x] **Parche cliente 001053** (limpiar PII: nombre, teléfonos, referencias, dirección, geo). Commit `acbf9e4`.
- [ ] **Crear hoja SinCastigoMarzo**. Aceptación: hoja idéntica a Marzo2026 pero sin registros con PAR=Mayor_180.
- [ ] **Mover créditos Oficina Central de Abril → Marzo**. Aceptación: registros con Coordinación="Oficina Central" salen de Abril2026 y aparecen en Marzo2026 (y en SinCastigoMarzo si cumplen filtro).
- [ ] **Agregar columna Fecha próximo pago**. Aceptación: nueva columna en R_Completo y hoja de fecha, cruzando con Rep_Cobranza por codigo+ciclo. Si fecha ≤ hoy → vacío. Requiere modificar UI para subir 2 archivos.

### Backlog técnico (pendiente de sprint actual)
- [ ] **Parametrizar fecha de corte mensual** (`Marzo2026`, `Abril2026`, `2026-04-01` hardcodeados). Bloquea corridas de meses distintos a abril 2026.
- [ ] **Criterio configurable de "Autorizado por Cartera"**. Hoy hardcodeado a `≠ Entregado`.
- [ ] **Crear `requirements.txt`**.
- [ ] **Mover `SECRET_KEY` a env vars**.
- [ ] **Sincronizar README**.

## Bloqueado / esperando
- **Flujo grupal**: STUB. Falta spec del dueño (columnas/hojas de salida, cruces entre 5 archivos).
- **Reporte de Juniors**: confirmado que es el input de Yunius (ReportedeAntiguedaddeCartera), no archivo separado.

## Hecho recientemente
- **`acbf9e4`** — Parches para cliente 001388 (Metepec) y 001053 (limpiar PII).
- **`ff779cb`** — Docs vivos `research.md` y `plan.md`.
- **`fe8d8e0`** — Hoja "Autorizado por Cartera".
- **`dcc07bb`** — Sistema `parches.json` + parche Oficina Central (18 acreditados).
