# PROJECT CHECKPOINT

Fecha: 2026-08-17
Proyecto: ERP-PYME-Manufactura (evolución de Propuesta.xlsm / NJ Store)

## FASE ACTUAL
Fase 0 — Descubrimiento de dominio, hecho de verdad sobre las fuentes primarias del negocio — **EN CURSO, no cerrada**.

(La Fase 0 marcada "completada" en el checkpoint anterior del 2026-08-16 no se hizo sobre las fuentes primarias, sino sobre un archivo intermedio — `RECETAS_NORMALIZADAS.xlsx` — que la dueña del negocio descartó por estar mal analizado. Esta sesión rehizo la Fase 0 de verdad. Ver `DECISIONS.md` D-007.)

## COMPLETADO
✓ Fase 0 real ejecutada sobre las cuatro fuentes primarias: `Propuesta.xlsm` (original, con motor LAMBDA completo), `propuesta-COTIZADOR.xlsx` (refactor anterior), `fRANELA DE mUSELINA Y mICRO.xlsx`, y auditoría completa (51/51 hojas) de `HOJA DE CALCULO DE PRECIOS AL MAYOR.xlsx`.
✓ Documento maestro: **`ANALISIS-DOMINIO-FASE0.md`** — léase primero en cualquier sesión nueva. Contiene 19 reglas de negocio confirmadas directamente con la dueña, los hallazgos de la auditoría completa, matriz de paridad original-vs-refactor, diccionario del negocio, y un modelo conceptual propuesto (sin código, sin SQL).
✓ Validado con la dueña del negocio (2026-08-17): ancho de tela variable por tipo de tela (no un valor fijo); categoría de consumo (Niño/Dama/Caballero/Plus/Unica/Accesorio) asignada por criterio propio en cada lote, no por tabla fija; recargo +$3 desde talla XL vigente (por mayor consumo real); abono del 60% vigente como protección contra cancelación tras iniciar producción (con excepción confirmada de 70% en pedidos grandes por encargo); tramos de margen ligados a desperdicio real de material, no solo política de precio; override ligado principalmente a variación de costo por color/inventario específico de avíos; ancho útil de vinil de 48cm confirmado (rollo de 50cm menos margen de rodillos del plotter).
✓ `RECETAS_NORMALIZADAS.xlsx` **DESCARTADO** por la propia dueña del negocio — ya no existe en la carpeta del proyecto y **no debe tratarse como fuente de verdad** en ningún trabajo futuro.
✓ `DECISIONS.md` actualizado con D-007: estado de D-001 a D-006 a la luz de esta Fase 0 real.

## HALLAZGO CRÍTICO DE ESTA SESIÓN
Las decisiones D-001 a D-006 se apoyaron en una Fase 0 previa hecha sobre una fuente que resultó descartada. D-004 (moneda configurable) queda confirmada y reforzada por la nueva evidencia. **D-001 (modelo de atributos genéricos configurables) queda señalada para reconsiderar** — la Fase 0 real no encontró evidencia suficiente para generalizar todavía; encontró en cambio un dominio textil concreto con una dimensión adicional no capturada antes (tipo/marca específica de tela) y al menos tres formas distintas de "talla" coexistiendo en el negocio. Detalle completo en `DECISIONS.md` D-007.

## PENDIENTE ANTES DE MODELAR O CONSTRUIR
- Cerrar con la dueña las 2 preguntas de negocio que siguen abiertas (`ANALISIS-DOMINIO-FASE0.md` sección 13): definir un corte único de tramos de volumen para el sistema nuevo (se encontraron 4 nomenclaturas distintas en el negocio, sin estándar); confirmar si el abono variable (60%/70%) es una regla consciente o varió caso a caso.
- Decidir explícitamente con la dueña si D-001 se mantiene, se pospone o se revierte (`DECISIONS.md` D-007).
- Cerrar el modelo de dominio concreto (catálogo + recetas + cotización — `ANALISIS-DOMINIO-FASE0.md` sección 14), incorporando los hallazgos de la auditoría completa: tercera dimensión de tela (tipo/marca específica), segunda línea de negocio (NJ Sport, con tarifas propias), tallado libre para pedidos por encargo.
- Solo después de lo anterior: `/opsx:propose` para el modelo de datos concreto.

## NO HACER TODAVÍA
- ❌ No pasar a `/opsx:propose`.
- ❌ No generar SQL, migraciones, componentes de frontend, fórmulas definitivas ni código de ningún tipo.
- ❌ No dar por cerrada la arquitectura de `ARCHITECTURE.md` sin revisar `DECISIONS.md` D-007 primero — su modelo de dominio depende de D-001, que está señalada para reconsiderar.
- ❌ No reutilizar `Cotizador_V1_365.xlsx` / `Cotizador_V1_365.2.xlsx` como punto de partida sin revisión — dependen de `RECETAS_NORMALIZADAS.xlsx`, que fue descartado.

## CAPACIDADES DEL ORIGINAL QUE SE DEBEN CONSERVAR OBLIGATORIAMENTE
Confirmadas como reales y perdidas en el refactor anterior (`ANALISIS-DOMINIO-FASE0.md` secciones 8 y 10):
1. **Override por línea de cotización** — en `Propuesta.xlsm` existe la columna "Unidad Manual", que reemplaza el costo calculado solo para esa línea sin tocar el catálogo maestro. Está ausente en `propuesta-COTIZADOR.xlsx` + `M01_GestionPresupuesto.bas`. Confirmado por la dueña como mecanismo de uso frecuente (variación de costo por color/inventario específico).
2. **Cálculo de anidado geométrico para consumo de vinil** — `ObtenerConsumoVinil` en `Propuesta.xlsm` calcula piezas por fila según el ancho útil del rollo (48cm) y filas necesarias según la cantidad pedida. Está ausente en el refactor, que no reproduce ese cálculo.

## PROBLEMAS ABIERTOS
Ninguno bloqueante para seguir en Fase 0. Bloqueante para pasar a Fase 1 (modelado): las 2 preguntas de negocio pendientes y la decisión sobre D-001 (ver arriba).

## ARCHIVOS RELEVANTES (actualizado 2026-08-17)
- **`ANALISIS-DOMINIO-FASE0.md`** — documento maestro de la Fase 0 real. Léase primero.
- `PROJECT.md`, `ARCHITECTURE.md`, `DECISIONS.md`, `DEFERRALS.md` — arquitectura y decisiones previas; `DECISIONS.md` D-007 documenta qué de esto sigue vigente, qué se reconsidera y qué queda en pausa.
- `Propuesta.xlsm`, `propuesta-COTIZADOR.xlsx`, `fRANELA DE mUSELINA Y mICRO.xlsx`, `HOJA DE CALCULO DE PRECIOS AL MAYOR.xlsx` — fuentes primarias, ya analizadas.
- `RECETAS_NORMALIZADAS.xlsx` — **no existe, descartado**. Si aparece en otra copia del proyecto, no usarlo como fuente.
- `Cotizador_V1_365.xlsx`, `Cotizador_V1_365.2.xlsx` — construidos sobre la fuente descartada; no confiar en ellos sin revisión completa contra `ANALISIS-DOMINIO-FASE0.md`.
