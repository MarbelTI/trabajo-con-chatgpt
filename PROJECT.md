# PROJECT.md

> **Estado (2026-08-17): ver `ANALISIS-DOMINIO-FASE0.md` primero.** Este archivo y `ARCHITECTURE.md` reflejan decisiones (D-001 a D-006) tomadas sobre una Fase 0 previa que se apoyó en `RECETAS_NORMALIZADAS.xlsx`, un archivo que la dueña del negocio descartó por estar mal analizado. Una Fase 0 real, hecha sobre las fuentes primarias del negocio, ya se completó y está documentada en `ANALISIS-DOMINIO-FASE0.md`. `DECISIONS.md` D-007 registra qué de lo que sigue en este archivo está confirmado, qué debe reconsiderarse (en particular D-001, más abajo en "Restricciones") y qué queda en pausa. No se debe construir sobre lo que sigue sin leer eso primero.

## Propósito
Producto de gestión (Excel + VBA) para PYME manufactureras que arman productos por configuración (talla, color, acabado, técnica, tamaño...). Reemplaza cotizaciones en hojas sueltas por un motor de costeo/presupuestos reutilizable, vendible como plantilla protegida a distintos negocios — no solo a NJ Store.

Punto de partida: `Propuesta.xlsm` (ERP interno de NJ Store, confección textil). Se reutiliza su lógica de negocio (costeo por actividad + material) pero se generaliza el modelo de datos para que no dependa de atributos textiles fijos.

## Hallazgo clave del archivo original
De los 8 módulos que aparecen en el menú "Inicio", solo 4 tienen datos y lógica reales:
- **Productos** — catálogo dimensional (34 SKU, categorías/subcategorías, variantes, tallas, diseño)
- **PresupuestoPlantilla** — cotizador con motor de costeo por Actividad (Act-XX) + Material (Mat-XX)
- **Materiales** — catálogo de 52 insumos con costo, moneda, proveedor
- **Finanzas** — comandas (pedidos), abonos, saldo

**Producción, Compras, Ventas, Clientes/CRM y Costos** son solo encabezados/placeholders sin tablas ni fórmulas. Esto no es un defecto a corregir de inmediato — es información que determina el orden de construcción (ver Alcance V1).

## Usuarios
- **Dueño/administrador del negocio comprador**: arma cotizaciones, ve costos reales, controla cobros.
- **Comprador del producto (nuevo cliente que instala la plantilla)**: configura su propio catálogo de productos/atributos/materiales al recibirla — no debe tocar fórmulas ni VBA.
- **Vendedor de la plantilla (tú)**: mantiene la versión maestra, genera copias protegidas por licencia.

## Problema que resuelve
ANTES (estado del archivo original): cotizador funcional pero acoplado 1:1 al negocio textil de NJ Store — atributos, actividades y materiales hardcodeados; sin protección; sin versión genérica vendible.
DESPUÉS: motor de costeo reutilizable por cualquier PYME manufacturera vía atributos configurables, protegido, con deck de venta propio.

## Criterio de éxito
- Un comprador sin conocimientos de Excel/VBA puede cargar su propio catálogo (productos, atributos, materiales, actividades) sin tocar fórmulas.
- El motor de costeo calcula correctamente para al menos 2 verticales distintas de prueba (ej. confección textil y otro rubro manufacturero) usando la misma arquitectura de tablas.
- Archivo protegido: hojas de motor y VBA bloqueados; el comprador solo edita zonas designadas.
- Deck de venta comunica el antes/después y el valor del motor de costeo a un dueño de taller sin conocimiento técnico.

## Restricciones (Marco)
- **Secuencia revisada (D-002)**: V1-365 primero (100% fórmulas, sin macros/VBA), V1-VBA después como puerto protegido para vender. Ver D-002. **En pausa (ver D-007):** la implementación previa de V1-365 (`Cotizador_V1_365.xlsx`) se construyó sobre `RECETAS_NORMALIZADAS.xlsx`, descartado — no reutilizar sin revisión.
- El usuario ya tiene instalado el complemento **Claude para Excel** — confirmado que no soporta macros ni VBA, por lo que encaja con la fase V1-365 y no con la fase VBA.
- Multi-moneda: el original hardcodea USD/VES/COP — se generaliza a tabla de monedas configurable. Ver D-004 (confirmada y reforzada por la Fase 0 real, ver D-007 — la tabla de tasas necesita vigencia por fecha, no un solo valor fijo).
- Protección (solo V1-VBA): contraseña + hojas bloqueadas + validación de licencia simple. Ver D-003.
- Entrega de VBA sigue `REGLAS-VBA.md`: archivos `.bas` que el usuario importa en su propio Excel — eso aplica a la fase V1-VBA, no a V1-365.
- Carpeta de trabajo local del usuario: `C:\Users\Asesor Comercial\Documents\PROYECTO TEXTIL` — Claude (este chat) no tiene acceso a esa ruta; los archivos se entregan aquí y el usuario los mueve/abre allí, o los trabaja en vivo con Claude para Excel una vez movidos.

## Alcance V1-365 (motor de valor, formulado, sin macros)
Construido sobre `RECETAS_NORMALIZADAS.xlsx` (dato real ya limpio, ver D-006): hoja **Cotizador** (selector + motor INDEX/MATCH/SUMIFS) + hoja **Indice** (236 configuraciones priceables) sobre `M_Recetas`/`M_Escalas`.
Exclusiones de V1: Producción, Compras, Ventas, Clientes/CRM, Inventario, Finanzas, catálogo de Atributos genérico (D-001 queda validado conceptualmente pero no implementado todavía — ver DEFERRALS D-005), versión VBA protegida — quedan en V1-VBA/V2/V3 (ver `DEFERRALS.md`).
