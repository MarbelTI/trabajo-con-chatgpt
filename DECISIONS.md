# DECISIONS.md

## D-001 — Generalización del catálogo (Atributos configurables)
Fecha: 2026-08-16
Decisión: reemplazar el modelo fijo Talla/Tela/Técnica del original por un modelo de Atributos configurables (`Dim_Atributo` + `Producto_Atributo`), tomando como base el patrón que ya existía en la hoja `Tabla4 (3)` del archivo original.
Motivo: el producto se vende a PYME manufactureras de cualquier rubro, no solo confección textil.
Impacto: todo el motor de costeo (Regla_Consumo) se diseña sobre atributos genéricos, no sobre campos textiles fijos.

## D-002 — Secuencia revisada: V1 formulado en 365 primero, motor VBA después (REVOCA la versión anterior de esta decisión)
Fecha: 2026-08-16 (actualizada el mismo día tras indicación explícita del usuario)
Decisión: se construye primero una versión **V1-365**, 100% fórmulas nativas (INDEX/MATCH/SUMIFS; XLOOKUP/LAMBDA opcional más adelante), sin macros ni VBA. La versión protegida en VBA (antes D-002 original) se construye **después**, como puerto de la lógica ya validada — no como V1.
Motivo: (1) el usuario lo pidió explícitamente; (2) el usuario ya tiene instalado el complemento **Claude para Excel**, que trabaja en vivo sobre el libro abierto pero **no soporta macros ni operaciones VBA** (confirmado en la documentación oficial de Anthropic) — encaja exactamente con una fase 100%-fórmulas y no con la fase VBA.
Impacto: el orden de construcción queda: **V1-365 (fórmulas) → validar con datos reales → V1-VBA (puerto protegido, para vender)**. La restricción de compatibilidad Excel 2016+ sigue aplicando *solo* a la versión VBA final que se vende; la versión 365 es una herramienta de trabajo/validación, no el entregable de venta.

## D-003 — Protección: VBA con clave + hojas bloqueadas
Fecha: 2026-08-16
Decisión: proteger el proyecto VBA con contraseña, bloquear hojas de motor/reglas, dejar editables solo las zonas de catálogo del comprador (Productos, Atributos, Materiales, Actividades).
Motivo: se vende como plantilla-producto ("caja negra"); el comprador no debe poder copiar el motor de cálculo.
Impacto: se requiere un checklist de qué queda editable vs. bloqueado antes de Fase 3 (`PONYTAIL-INTEGRATION.md` + revisión de zonas de entrada).
Pendiente de definir: mecanismo de validación de licencia por comprador (nivel de complejidad a decidir en Fase 1 de detalle si aplica).

## D-004 — Moneda configurable en vez de USD/VES/COP hardcodeado
Fecha: 2026-08-16
Decisión: reemplazar las columnas de moneda fijas del original por una tabla `Moneda` (código, tasa vs. moneda base) configurable por el comprador.
Motivo: cada negocio comprador opera en su propio país/moneda.
Impacto: todas las fórmulas de costeo referencian moneda base + tasa, no un valor fijo.

## D-006 — Fuente de datos real para el motor de costeo: RECETAS_NORMALIZADAS.xlsx
Fecha: 2026-08-16
Decisión: el motor de costeo V1 se construye directamente sobre `RECETAS_NORMALIZADAS.xlsx` (trabajo de limpieza ya hecho por el usuario en un proyecto anterior), no sobre datos de ejemplo inventados.
Hallazgo: el archivo contiene `M_Recetas` (2,827 líneas de componente, 236 configuraciones/`id_receta`, 32 familias de producto), `M_Escalas` (precio por escala de volumen: +5 DOC / 1 A 5 DOC / DETAL), `M_Insumos`, `M_Procesos` y `Anomalias` (auditoría propia de inconsistencias: tasas de cambio incrustadas distintas, mismo insumo con precios muy variables por usar nombres genéricos como "TELA" en vez de un código específico).
Impacto: `M_Recetas` ya cumple el rol de `Regla_Consumo` del modelo de dominio (trae consumo, precio unitario y costo ya calculados por línea) y `M_Escalas` el de la regla de precio por volumen — no hace falta reconstruirlas desde cero. `bloque_variante` se usa en V1 como clave de configuración ya armada (opaca), sin descomponerla todavía en atributos atómicos (Talla/Tela/Técnica) — esa descomposición queda en `DEFERRALS.md` para cuando se necesite reutilizar el motor en otro rubro.
Riesgo heredado (no resuelto, no bloqueante para V1): ~21 líneas de `M_Escalas` tienen una escala no estándar (`?`, `2 A 5 DOC`, `3 A 5 DOC`) y no devuelven precio en el cotizador — se muestran con aviso explícito en vez de fallar en silencio.

## D-005 — Orden de trabajo: deck de venta antes de construir todos los módulos
Fecha: 2026-08-16
Decisión: construir primero un deck de venta (concepto + antes/después + arquitectura V1) para validar interés antes de invertir sesiones completas en Producción/Compras/Ventas/CRM (módulos que en el original ni siquiera tenían datos).
Motivo: evitar sobreconstrucción (Ponytail Check) de módulos que podrían no ser el gancho de venta real. El motor de costeo (Productos + Cotizador) es el diferenciador — se valida y se construye primero.
Impacto: V1 técnico = Productos + Cotizador. Producción/Compras/Ventas/CRM se diseñan solo si el deck valida interés o el comprador los pide explícitamente.

## D-007 — Estado de D-001 a D-006 tras la Fase 0 real (2026-08-17)
Fecha: 2026-08-17
Contexto: D-001 a D-006 se tomaron apoyadas en una Fase 0 anterior que no se hizo sobre las fuentes primarias del negocio, sino sobre `RECETAS_NORMALIZADAS.xlsx` (ver D-006 arriba). Ese archivo fue **descartado por la propia dueña del negocio** ("estaba mal analizado") y ya no existe en la carpeta del proyecto. En esta sesión se ejecutó la Fase 0 real sobre las cuatro fuentes primarias (`Propuesta.xlsm`, `propuesta-COTIZADOR.xlsx`, `fRANELA DE mUSELINA Y mICRO.xlsx`, y auditoría completa de las 51 hojas de `HOJA DE CALCULO DE PRECIOS AL MAYOR.xlsx`), con 19 reglas de negocio validadas directamente con la dueña. Documento completo: `ANALISIS-DOMINIO-FASE0.md`.

Estado de cada decisión anterior a la luz de esa Fase 0 real:

- **D-001 (atributos genéricos configurables) — A RECONSIDERAR.** La Fase 0 real no encontró evidencia suficiente en las fuentes primarias para justificar generalizar ya a un modelo de atributos abstractos. Encontró en cambio un dominio textil concreto con reglas propias: el consumo de tela depende de producto × categoría de consumo × **tipo/marca específica de tela** (una tercera dimensión que no estaba capturada en ningún archivo previo, incluido el modelo de `Tabla4 (3)`/`Consumo_Tela_Producto`); y la "talla" tiene al menos tres formas distintas de aparecer en el negocio (categoría de consumo Niño/Dama/Caballero/Plus/Unica/Accesorio para el costeo; talla comercial S/M/L/XL/2XL para el recargo de precio; talla libre del cliente para pedidos grandes por encargo). Generalizar sobre esto antes de cerrarlo es el mismo riesgo que ya corrió esta decisión, apoyada entonces en una fuente que resultó descartada. No se revoca formalmente aquí — queda señalada para revisión explícita con la dueña antes de modelar (ver `ANALISIS-DOMINIO-FASE0.md` secciones 14 y 17).
- **D-002 (V1-365 primero, VBA después) — secuencia no contradicha, implementación previa en pausa.** El orden en sí (fórmulas antes de VBA) no fue contradicho por la Fase 0. Pero `Cotizador_V1_365.xlsx` / `Cotizador_V1_365.2.xlsx` se construyeron sobre `RECETAS_NORMALIZADAS.xlsx`, ya descartado — no deben tratarse como punto de partida sin revisión completa. La secuencia queda en pausa hasta cerrar el modelo de dominio real.
- **D-003 (protección VBA con clave + hojas bloqueadas) — sin cambios.** No afectada por la Fase 0 (aplica a una fase de construcción posterior). Queda en pausa junto con toda la construcción.
- **D-004 (moneda configurable, tabla de tasas) — CONFIRMADA y reforzada.** La Fase 0 real encontró exactamente el problema que esta decisión buscaba resolver: la tasa de cambio no está centralizada en ninguna de las fuentes manuales del negocio (referenciada en unas fórmulas, escrita literal en otras, con valores distintos incluso dentro del mismo archivo — ver `ANALISIS-DOMINIO-FASE0.md` sección 3.2). Observación que aporta la evidencia, para cuando se implemente D-004: la tabla de tasas necesita vigencia por fecha, no un solo valor fijo, porque las fuentes muestran tasas distintas en momentos distintos del negocio.
- **D-005 (deck de venta antes de construir todos los módulos) — sin cambios.** No afectada por la Fase 0 (es orden de negocio/validación de venta, no de datos).
- **D-006 (`RECETAS_NORMALIZADAS.xlsx` como fuente real del motor de costeo) — DESCARTADA explícitamente por la dueña del negocio** (2026-08-17: "estaba mal analizado", archivo eliminado de la carpeta del proyecto). Reemplazada como fuente de verdad por `Propuesta.xlsm`, `propuesta-COTIZADOR.xlsx`, `fRANELA DE mUSELINA Y mICRO.xlsx` y `HOJA DE CALCULO DE PRECIOS AL MAYOR.xlsx`, analizados en `ANALISIS-DOMINIO-FASE0.md`. Cualquier trabajo que dependa de `M_Recetas`/`M_Escalas`/`M_Insumos`/`M_Procesos`/`Anomalias` de ese archivo (incluido `Cotizador_V1_365.xlsx`) debe revisarse contra la nueva fuente antes de reutilizarse.

Impacto: `ARCHITECTURE.md` no se reescribió en esta sesión — su modelo de dominio (basado en D-001) puede estar desactualizado. No debe tratarse como cerrado hasta que D-001 se resuelva explícitamente con la dueña del negocio.
