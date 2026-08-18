# DEFERRALS.md

## D-001
Funcionalidad: Módulo Materiales/Inventario avanzado (stock, entradas/salidas, valorización).
Razón: no bloquea la venta del motor de costeo (V1). En el original ya existe un catálogo base de Materiales que se reutiliza como dato maestro simple en V1; el control de inventario completo es V2.
Estado: aplazada.
Objetivo: V2.

## D-002
Funcionalidad: Módulo Finanzas (flujo de caja, comandas, abonos como módulo independiente con reportes).
Razón: V1 incluye abonos/saldo a nivel de cotización individual (suficiente para vender); el flujo de caja consolidado es V2.
Estado: aplazada.
Objetivo: V2.

## D-003
Funcionalidad: Producción, Compras, Ventas, Clientes/CRM.
Razón: en el archivo original estos módulos eran solo encabezados sin tablas ni lógica — no hay nada que generalizar todavía, y construirlos sin validar el motor de costeo primero es sobreconstrucción.
Estado: aplazada.
Objetivo: V3, solo si el deck de venta valida interés o un comprador específico los pide.

## D-004
Funcionalidad: Validación de licencia por comprador (activación/control de copias).
Razón: depende de decidir el canal de venta real (venta directa 1 a 1 definida; marketplace descartado por ahora) — nivel de protección puede ser más simple de lo inicialmente previsto.
Estado: aplazada, a definir en Fase 1 de detalle antes de construir el módulo de protección.
Objetivo: V1-VBA, antes de la primera venta (no aplica a V1-365, que es herramienta interna de validación).

## D-005
Funcionalidad: Descomponer `bloque_variante` en atributos atómicos (Talla/Tela/Técnica/Segmento como campos separados, en vez de una etiqueta de texto).
Razón: para V1-365 la etiqueta ya identifica una configuración priceable de forma inequívoca (vía `id_receta`); descomponerla ahora sin un caso de uso concreto que lo exija es sobreconstrucción.
Estado: aplazada.
Objetivo: cuando se necesite reutilizar el motor para un rubro distinto a confección (validación real del D-001 de generalización), o si el comprador necesita filtrar/reportar por atributo individual.

## D-006
Funcionalidad: Resolver las ~21 líneas de `M_Escalas` con escala no estándar (`?`, `2 A 5 DOC`, `3 A 5 DOC`) y las inconsistencias de tasa de cambio listadas en `Anomalias`.
Razón: no bloquea el uso del cotizador — esas combinaciones puntuales muestran aviso en vez de un precio, en lugar de fallar en silencio o inventar un valor.
Estado: aplazada.
Objetivo: antes de considerar V1-365 "cerrado" para uso diario sin supervisión.

**Nota (2026-08-17):** `M_Escalas` y `Anomalias` pertenecían a `RECETAS_NORMALIZADAS.xlsx`, descartado por la dueña del negocio (ver `DECISIONS.md` D-007). Este ítem queda superado por los hallazgos de la Fase 0 real, hecha sobre fuentes vigentes: ver `ANALISIS-DOMINIO-FASE0.md` sección 9 (inconsistencias del original) y sección 16.7 (inconsistencias nuevas de la auditoría completa de `HOJA DE CALCULO DE PRECIOS AL MAYOR.xlsx`), incluyendo el hallazgo de que existen **al menos 4 nomenclaturas distintas de tramos de volumen** en el negocio, sin un corte único definido — pregunta de negocio todavía abierta (`ANALISIS-DOMINIO-FASE0.md` sección 13).
