# ARCHITECTURE.md

> **Estado (2026-08-17):** el modelo de dominio de abajo se basa en D-001 (atributos genéricos configurables), decisión que `DECISIONS.md` D-007 señala **para reconsiderar** a la luz de una Fase 0 real hecha sobre las fuentes primarias del negocio (ver `ANALISIS-DOMINIO-FASE0.md`). Este archivo no se reescribió en esa sesión — no tratar el modelo de abajo como cerrado hasta que D-001 se resuelva explícitamente.

## Modelo de dominio

Pregunta de dominio: ¿qué determina que un producto/pedido sea diferente de otro?
Respuesta (generalizada desde el original): **la combinación de atributos configurables que el negocio comprador define para su propio catálogo** — no una lista fija de Talla/Tela/Técnica.

El propio archivo original ya insinuaba esto en la hoja `Tabla4 (3)` (`PRODUCTO | Atributo | Valor`) — se adopta esa forma como base del motor, en vez del modelo rígido `Producto → Talla → Tela → Técnica → Material → Diseño → Modelo`.

```
ENTRADA (selección de producto + atributos del pedido)
   ↓
MODELO
   Dim_Producto        (código, nombre, categoría, subcategoría)
   Dim_Atributo         (id, nombre, tipo) — ej. Talla, Color, Tamaño, Acabado; lo define el comprador
   Producto_Atributo    (qué atributos aplican a qué producto + valores válidos)
   Dim_Material         (código, nombre, costo unitario, moneda, proveedor, stock)
   Dim_Actividad        (código, nombre, costo por unidad de operación)
   Dim_Cliente
   ↓
REGLAS
   Regla_Consumo         (Producto + Atributo → Material/Actividad → cantidad consumida)
   Recargo_por_Atributo   (ej. "+$3 desde talla XL" del original → parametrizado, no texto libre)
   Moneda_Base + Tasas    (generaliza el USD/VES/COP hardcodeado)
   % Abono mínimo
   ↓
MOTOR
   UDF en VBA: CalcularCosteo(producto, atributos[]) → recorre Regla_Consumo,
   suma (Material × costo) + (Actividad × costo)
   Reemplaza las LAMBDA/LET/XLOOKUP del original (no compatibles fuera de M365) — ver D-002.
   ↓
RESULTADO
   Cotización/Comanda (desglose, subtotal, total, abono, saldo) — misma lógica que
   `PresupuestoPlantilla` + `Finanzas` del original, generalizada a cualquier vertical.
```

Clasificación:
- **Datos maestros**: Productos, Atributos, Materiales, Actividades, Clientes
- **Transacciones**: Comandas/Cotizaciones, Abonos
- **Configuraciones**: combinación de atributos por línea de cotización
- **Reglas de negocio**: consumo por atributo, recargo por atributo, moneda base, % abono mínimo
- **Cálculos**: costo unitario, precio con margen, total, saldo pendiente
- **Catálogos**: monedas, unidades de medida
- **Resultados derivados**: cotización imprimible, comprobante de abono

## Módulos V1

### Módulo Productos
Tablas: `Dim_Producto`, `Dim_Atributo`, `Producto_Atributo`.
Flujo: alta de producto → asignar atributos aplicables → asignar valores válidos por atributo.
Salida: catálogo consultable desde el cotizador.

### Módulo Presupuestos/Cotizador
Tablas: `Dim_Material`, `Dim_Actividad`, `Regla_Consumo`, `Recargo_por_Atributo`, `Dim_Cliente`, `Comanda`, `Comanda_Detalle`, `Abono`.
Flujo: seleccionar producto → elegir atributos del pedido → motor calcula consumo y costo → genera cotización con subtotal/total/abono/saldo (igual que el original, pero orientado por reglas, no hardcodeado).
Formularios: selección en cascada producto → atributos aplicables (evita combinaciones inválidas).

## V1-365 — implementación real (ya construida)

Sobre `RECETAS_NORMALIZADAS.xlsx`, sin tocar `M_Recetas`/`M_Escalas`/`M_Insumos`/`M_Procesos`/`Anomalias` (quedan como datos fuente):

- **Indice** (nueva): 236 filas, una por `id_receta`, con etiqueta legible `hoja — bloque_variante [#id]` para el desplegable.
- **Cotizador** (nueva): 3 entradas (Configuración, Escala, Cantidad) → motor:
  - Costo unitario = `SUMIFS(M_Recetas!costo_usd, id_receta=seleccionado)`
  - Precio unitario = `INDEX/MATCH(M_Escalas!precio_usd, clave=id_receta&"|"&escala)`
  - Desglose por categoría (Insumo/Proceso/Tecnica/Indirecto/Otro) con `SUMIFS` de dos criterios, con un total de verificación que debe igualar el costo unitario.
- Solo `INDEX`/`MATCH`/`SUMIFS` (Excel 2007+, funciona igual en 365) — se evitó `XLOOKUP`/`LAMBDA` en esta build porque el entorno de verificación (LibreOffice) no evalúa correctamente funciones de array dinámico, y esta versión sí se pudo probar con `recalc.py` contra datos reales. Si se quiere sintaxis más "365-nativa" (`XLOOKUP`/`LAMBDA`), se puede pedir directamente a **Claude para Excel** que la reescriba en vivo sobre el archivo abierto — ese complemento sí valida contra Excel real.
- `bloque_variante` se usa como clave opaca vía `id_receta` (ver DEFERRALS D-005) — no se descompuso en atributos atómicos todavía.
- Casos sin precio (escala `?`, `2 A 5 DOC`, `3 A 5 DOC` — ver `Anomalias`) muestran aviso explícito en vez de fallar en silencio o inventar un valor.

## Diferido a V2/V3
Ver `DEFERRALS.md` — Materiales/Inventario avanzado, Finanzas (flujo de caja), Producción, Compras, Ventas, Clientes/CRM.

## Permisos
V1: usuario único (dueño del negocio comprador). Multiusuario queda fuera de alcance V1.

## Dependencias
Ninguna externa. Todo el motor vive en el archivo `.xlsm` (VBA + tablas), sin conexión a servicios externos.
