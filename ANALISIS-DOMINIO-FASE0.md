# Análisis de dominio — Fase 0 (arqueología del cotizador textil)

## Estado para continuar en otra sesión/cuenta (2026-08-17)

Léase esto primero. Ver también `PROJECT-CHECKPOINT.md` (checkpoint corto) y `DECISIONS.md` D-007 (estado decisión por decisión).

1. **Reglas de negocio confirmadas**: 19 reglas, todas validadas con la dueña del negocio o verificadas por fórmula/texto directo en los archivos. Lista completa en la sección 6 ("Confirmadas").
2. **Hallazgos nuevos de la auditoría completa** (51/51 hojas de `HOJA DE CALCULO DE PRECIOS AL MAYOR.xlsx`): sección 16 — incluye la resolución de los 7 casos de consumo ambiguo, una tercera dimensión necesaria en el modelo de tela (tipo/marca específica), una segunda línea de negocio (NJ Sport) con tarifas propias, tallado libre para pedidos por encargo, abono variable (60%/70%), y 4 inconsistencias de fórmula nuevas.
3. **Decisiones de arquitectura anteriores (D-001 a D-006)**: su estado actualizado —confirmada, a reconsiderar, o en pausa— está en `DECISIONS.md` D-007, no en este documento. Resumen: D-004 confirmada y reforzada; D-001 señalada para reconsiderar (no hay evidencia suficiente todavía para generalizar a atributos genéricos); D-002/D-003 en pausa; D-005 sin cambios; D-006 original **descartada** (ver punto 6).
4. **Preguntas de negocio todavía abiertas** (solo 2, de las 10 originales): sección 13, bloque "Resueltas por la auditoría completa" explica qué se cerró y qué sigue abierto — el corte único de tramos de volumen para el sistema nuevo, y si el abono 60%/70% es una regla consciente.
5. **Capacidades del original que se deben conservar obligatoriamente**: override por línea de cotización (columna "Unidad Manual", sección 8) y cálculo de anidado geométrico de vinil (`ObtenerConsumoVinil`, sección 3.3) — ambas confirmadas como reales y ausentes en el refactor (`propuesta-COTIZADOR.xlsx` + `M01_GestionPresupuesto.bas`), ver también sección 10.
6. **`RECETAS_NORMALIZADAS.xlsx` fue descartado** por la dueña del negocio ("estaba mal analizado") y ya no existe en la carpeta. No es fuente de verdad. La fuente real para todo el motor de costeo es: `Propuesta.xlsm` + `propuesta-COTIZADOR.xlsx` + `fRANELA DE mUSELINA Y mICRO.xlsx` + `HOJA DE CALCULO DE PRECIOS AL MAYOR.xlsx` — las cuatro ya analizadas en este documento. Ver sección 0.
7. **La auditoría de las 51 hojas ya se hizo** y sus conclusiones están en la sección 16 de este mismo documento — no repetirla desde cero.
8. **Todavía NO se debe pasar a `/opsx:propose` ni empezar implementación** — falta cerrar las 2 preguntas del punto 4 y decidir explícitamente el estado de D-001 con la dueña del negocio. Ver sección 17 (Recomendación de siguiente paso).

---

Convención usada en todo el documento:
- **[DICE]** — está escrito literalmente en el archivo (celda, fórmula o texto).
- **[INFIERE]** — se deduce con buena confianza cruzando varias evidencias, pero no está escrito explícitamente.
- **[HIPÓTESIS]** — posible explicación, sin evidencia suficiente para confirmarla.
- **[PROPONGO]** — diseño que yo sugiero para el sistema nuevo. No es un hecho del negocio.

---

## 0. Aviso — tensión que no se resuelve en silencio

El archivo 3 fue descrito como "documentación/arquitectura existente del proyecto". **No lo es.** `fRANELA DE mUSELINA Y mICRO.xlsx` es una hoja de costeo manual (315 filas) con las mismas fórmulas de consumo de tela que el archivo original — un cuarto conjunto de datos crudos del mismo negocio, no una decisión de arquitectura. La documentación de arquitectura real está en `ARCHITECTURE.md` / `DECISIONS.md` / `PROJECT.md` / `DEFERRALS.md`, que sí existen en la carpeta pero no fueron uno de los tres archivos entregados.

Además, esos documentos revelan algo que cambia el marco de esta tarea: **ya hubo una Fase 0 previa** (`PROJECT-CHECKPOINT.yaml`: `current_phase: "0-Descubrimiento ROMA... (completado)"`, 2026-08-16) que terminó en decisiones de arquitectura (D-001 a D-006) y en un cotizador V1-365 ya construido y marcado "verificado" (`Cotizador_V1_365.xlsx`). Pero esa Fase 0 **no se hizo sobre `Propuesta.xlsm` directamente** — se hizo sobre un archivo intermedio, `RECETAS_NORMALIZADAS.xlsx` (D-006), que **no existe en esta carpeta ni en ningún subdirectorio** (confirmado por búsqueda recursiva). Es decir: la arquitectura ya decidida (D-001: modelo de atributos genéricos; D-004: moneda configurable) se apoyó en un archivo de trabajo que ya no está disponible para auditar, no en los tres archivos que ahora tengo delante.

Esto significa que este análisis es, en la práctica, **la Fase 0 que debió hacerse antes de D-001–D-006**, no una repetición cosmética. Donde el análisis de abajo contradiga esas decisiones, lo digo explícitamente en la sección 10 y en las preguntas críticas — no las doy por buenas solo porque ya están escritas en `DECISIONS.md`.

**Actualización (misma fecha, tras validación con la dueña del negocio)**: `RECETAS_NORMALIZADAS.xlsx` fue eliminado de la carpeta por la propia dueña — "estaba mal analizado". La fuente real de la que debió salir ese archivo es `HOJA DE CALCULO DE PRECIOS AL MAYOR.xlsx` (377 KB, **51 hojas, una por línea de producto**: BUZO, SUETER, FRANELILLAS, SHORT, LICRAS, UNI FUTBOL, CAMISAS, etc. — `FRA MUSE MICR` es la misma hoja que el archivo 3 original, confirmando que ese archivo 3 era un extracto de este libro maestro, no un documento aparte). Es sustancialmente más grande que los tres archivos originales: mismo patrón de fórmulas (`consumo × precio_tela × tasa`, tramos `+5 DOC/1 A 5 DOC/DETAL`, bloques por tipo de tela) pero con el catálogo completo del negocio, no solo franelas. Alcance de esta actualización: se incorporan las respuestas de la dueña a las preguntas críticas y una revisión dirigida (no exhaustiva) de este archivo nuevo — una auditoría completa de las 51 hojas es un trabajo de escala distinta, pendiente de decidir si se hace (ver cierre del documento).

---

## 1. Resumen ejecutivo

`Propuesta.xlsm` no es un Excel simple: tiene un **motor de costeo real construido con LAMBDA/LET/FILTER/XLOOKUP** (11 funciones nombradas: `obtenerCostoTela`, `obtenerCostoSublimado`, `obtenerCostoVinil`, `ObtenerCostoDTF`, `obtenerCostoMaterial`, `obtenerCostoActividad`, `obtenerCostoGlobal` como despachador, más sus `obtenerConsumo*`). Este motor ya expresa, en código, casi todo lo que el negocio necesita: consumo de tela por producto+talla, consumo de vinil por área de diseño y cantidad (con lógica de acomodo/anidado sobre el ancho del rollo), consumo de sublimado derivado del tamaño de estampado o de la tela, costo de actividades por producto y por tramo de volumen, con conversión de moneda centralizada vía un parámetro `tasa_dolar` que se pasa explícitamente por toda la cadena de funciones.

Sin embargo, ese motor **descansa sobre datos crudos con huecos reales**: ~20% de las filas de la tabla maestra de consumo de tela (`Tabla4 (3)`) tienen **dos valores en la misma celda** ("0,9 Y 1,1", "0,6/2; 0,8"), sin resolver cuál aplica; dos filas usan el nombre del producto en vez de su código como llave; y una función completa (`CostoMaterialDecorativo`) es un *stub* sin terminar que devuelve el texto `"Prueba2"`. El motor es más sofisticado que sus datos.

El cotizador refactorizado (`propuesta-COTIZADOR.xlsx`) **rescató correctamente** la estructura relacional (tablas de Productos/Materiales/Producción separadas, tasas de cambio centralizadas en su propia hoja) pero **perdió una capacidad de negocio real y documentada del original**: la columna "Unidad Manual" que permitía anular el costo calculado de una línea sin tocar el maestro (ver sección 8). El refactor tampoco reprodujo el motor de consumo de vinil por área/anidado ni el recargo de +$3 desde talla XL que el propio original documenta en texto plano dentro de la hoja.

La tercera fuente (`fRANELA DE mUSELINA Y mICRO.xlsx`) no es arquitectura: es la prueba de campo de la misma fórmula de costeo de tela (consumo × precio USD/metro × tasa de cambio) replicada a mano decenas de veces, con la tasa de cambio unas veces centralizada en una celda y otras veces escrita literal en la fórmula — exactamente la inconsistencia que se pidió investigar.

---

## 2. Qué hace realmente el cotizador original

`Propuesta.xlsm` es el ERP interno de un taller de confección ("NJ Store", texto literal en `PresupuestoPlantilla!B19`). De sus 18 hojas, **8 tienen datos y lógica real**: `PresupuestoPlantilla` (el cotizador), `Productos`, `Materiales`, `Producción`, `Auxiliares`, `Tabla4 (3)` (consumo de tela), `extraccion manual` (staging de datos crudos), `Finanzas`. El resto (`Compras`, `Ventas`, `Clientes _CRM`, `Costos`, `Costos Adicionales`) están vacías u ocultas — confirma lo que ya había notado la sesión anterior en `PROJECT.md`.

El flujo real, reconstruido desde las fórmulas (`PresupuestoPlantilla`, fila 5-7 y celdas K1/K6/K7):

1. El operador elige, en celdas de entrada: Producto, Talla, Tela, Técnica, Tamaño (según técnica), Material, Cantidad.
2. Fórmulas `XLOOKUP` resuelven cada nombre elegido a su código (`L7`, `M7`, `N7`, `O7`... contra `Dim_Producto`, `Dim_Talla`, `HechosInventario`, `dim_Tecnica`).
3. `K7 = IF(K6<=0,"S/N",IF(K6<12,"VC001",IF(K6<1000,"VC002","VC003")))` clasifica la cantidad en un **tramo de volumen** (Detal / +12 / +1000).
4. Cada línea de costo (`Costro_Producto`, tabla con columnas Codigo|Concepto|Unidad|Unidad Manual|Sub Total|Total) resuelve su costo unitario llamando a `obtenerCostoGlobal`, que despacha según el código: `Act-11`→vinil, `Act-10`→sublimado, `Act-12`→DTF, prefijo `ACT`→costo de actividad por tabla, prefijo `MAT`→costo de material.
5. `Sub Total × K6 (cantidad)` = `Total` por línea; `SUBTOTAL(109, Costro_Producto[Total])` = total del pedido.
6. El total en pesos se divide por `K1` (tasa dólar) para obtener el **costo unitario en USD** (`O29 = Costro_Producto[[#Totals],[Sub Total]]/K1`) — el negocio calcula en COP línea por línea y **normaliza a USD al final**, no al revés.
7. `Finanzas` registra Comanda (pedido) + Abono + Saldo — la cotización aceptada se convierte en comanda, con abono parcial (60% según `B24`, ver sección 9).

---

## 3. Anatomía de los cálculos

### 3.1 Costo de tela — la fórmula central

```
obtenerCostoTela(producto, talla, cod_material, tasa_dolar):
    consumo       = XLOOKUP en Consumo_Tela_Producto donde Codigo_Producto=producto Y Talla=talla → Cantidad_Tela_m²
    precio        = XLOOKUP del material → Costo_Unitario_Detal
    moneda        = XLOOKUP del material → Moneda
    precio_final  = SI moneda="USD" ENTONCES precio × tasa_dolar SINO precio
    resultado     = consumo × precio_final
```
**[DICE]**, tomado directamente de la definición LAMBDA en `Propuesta.xlsm`.

La forma "artesanal" de la misma fórmula, encontrada literal en `fRANELA...xlsx!B4`: `=0.8*2.3*E2`, y en decenas de celdas más del mismo archivo. Descomposición verificada por consistencia aritmética (no solo supuesta):

- **0.8** = consumo de tela en metros para esa combinación producto+categoría de talla — mismo rol que `Cantidad_Tela_m²` en el motor formal. **[INFIERE]**, por comparación directa: los mismos valores (0.6, 0.8, 0.9, 1.0...) aparecen en `Tabla4 (3)!C` del archivo original como `Cantidad_Tela_m²`, y en `propuesta-COTIZADOR.xlsx!Produccion!AM` como `Cantidad_Tela_m²`. Tres fuentes independientes coinciden en la magnitud y en el rol.
- **2.3** = precio de la tela en USD por metro. **[INFIERE]** con alta confianza: en `propuesta-COTIZADOR.xlsx!Materiales`, `Tela Muselina` tiene `Costo_Unitario_Detal = 2.3 USD` (fila 34) — coincide exactamente con el 2.3 usado en `fRANELA...!B4` para "MUSELINA DAMA".
- **E2 (=4000)** = tasa de cambio COP por USD, no un precio de tela ni una cantidad de producción. **[INFIERE]**, confirmado por consistencia dimensional: `fRANELA...!B12 = SUM(B4:B11)/E2 = 3.6525` — sumar costos en pesos (tela+corte+costura+etiqueta+bolsa+marquilla+hilo+gastos) y dividir por 4000 solo da un número con sentido de negocio (~USD 3.65, costo de producción de una franela en blanco) si E2 es la tasa de cambio. Es el mismo patrón, verificado independientemente, que `PresupuestoPlantilla!O29 = Total/K1(=4000)`.

Es decir: **consumo(m) × precio_USD/m × tasa_cambio = costo_tela_en_COP**, y luego, al final del escandallo completo, **costo_total_COP / tasa_cambio = costo_unitario_en_USD**. El negocio calcula operativamente en pesos línea por línea, y presenta/decide en dólares. Esto aparece confirmado en tres archivos independientes (motor LAMBDA, hoja manual de franelas, hoja de presupuesto), así que lo trato como **regla confirmada**, no hipótesis.

### 3.2 La tasa de cambio no está centralizada — está repetida

Lo que el usuario pidió verificar explícitamente. Evidencia concreta de ambos comportamientos **en el mismo archivo**:

- **Si se referencia**: `fRANELA...!B4 = 0.8*2.3*E2`, `!B14 = 0.3*1.3*E2`, decenas más — todas dentro del bloque "superior" del archivo referencian `E2`.
- **Si se hardcodea**: `fRANELA...!G111 = 1*1.2*3800` (no usa E2 ni E140, escribe 3800 literal); `!H9 = H8*3800`; `!C117 = C116*3700`; `!D117 = D116*3700`; `!H181 = H180*3800`; `!D197 = D196*3800`. Los valores hardcodeados (3700, 3800, 3900) **no coinciden entre sí ni con los valores de las celdas de referencia** (E2=4000, E140=4500) — son tasas de distintos momentos, nunca reconciliadas.
- Peor: el mismo archivo tiene **dos celdas de referencia distintas** (`E2=4000` y `E140=4500`) usadas cada una en su zona del archivo, y bloques visualmente lejanos de `E140` (fila 266, 284, 301) la siguen referenciando en vez de crear una tercera — es decir, no hay una sola fuente de verdad ni siquiera cuando SÍ se centraliza.
- En `Propuesta.xlsm`, el motor LAMBDA sí hace lo correcto: `tasa_dolar` es un **parámetro explícito** que se pasa por toda la cadena de funciones (`obtenerCostoTela(producto,talla,cod_material,tasa_dolar)`, etc.), alimentado desde `PresupuestoPlantilla!K1`. Este es el diseño más limpio de los tres archivos — pero solo aplica al motor nuevo (LAMBDA), no a las hojas de trabajo manual como `fRANELA...xlsx`, que es anterior o paralela y nunca se migró a usar ese parámetro.

**Conclusión**: el intento de centralizar la tasa **sí existe y en la versión LAMBDA está bien resuelto**, pero conviven en el negocio otras hojas de cálculo manual (como el archivo 3) donde la tasa se escribe a mano, se copia de una celda cercana, o ambas cosas a la vez sin ningún criterio visible de cuándo usar cada una.

### 3.3 Consumo de vinil — no es un número fijo, es un cálculo de anidado

```
ObtenerConsumoVinil(cod_tamano, cantidad):
    ancho, alto        = XLOOKUP en Dim_Tecnica_Tamano por Id_Tamano
    ancho_util         = 48                                    ← hardcodeado, probablemente cm útiles del rollo de vinil
    piezas_por_fila    = INT(ancho_util / ancho)
    filas_necesarias   = ROUNDUP(cantidad / piezas_por_fila, 0)
    alto_total         = filas_necesarias × alto
    consumo            = (ancho_util × alto_total) / 10000     ← cm² a m²
```
**[DICE]**. Esto es un cálculo real de aprovechamiento de material (cuántas piezas de tamaño ancho×alto caben en una fila de un rollo de 48 cm útiles, cuántas filas hacen falta para la cantidad pedida). No es ruido ni un intento fallido — es lógica de negocio legítima y más sofisticada que un simple "área × cantidad". Las hojas manuales (`fRANELA...`) nunca reproducen este cálculo: ahí el costo de vinil aparece como valores flat divididos entre un rendimiento fijo (ej. `B92 = 30000/39`, "una lámina de $30.000 rinde 39 diseños" — ver 3.4), un enfoque más simple que no considera el acomodo geométrico.

### 3.4 Consumibles por rendimiento — patrón distinto al de la tela

En `fRANELA...xlsx` aparecen fórmulas como `B89 = 13000/15`, `B92 = 30000/39`, `G191 = 15000/4`, `G206 = 35000/4`. **[INFIERE]**: esto es "costo de una lámina/rollo de insumo (papel de sublimación, vinil) dividido entre cuántas unidades rinde", un patrón de costeo por rendimiento — **distinto** del patrón tela (consumo×precio×tasa) y **distinto** del patrón vinil por área del motor LAMBDA. Es un tercer patrón de cálculo de consumo, válido y consistente en sí mismo, que el modelo futuro debe soportar además de los otros dos.

---

## 4. Modelo de productos y consumos

**[DICE + INFIERE, con alta confianza]** — confirmado cruzando `Propuesta.xlsm!extraccion manual` (fila 4, XLOOKUP contra `Dim_Talla`), `Tabla4 (3)` y `propuesta-COTIZADOR.xlsx!Productos`:

```
Tall-01 = Niño
Tall-02 = Dama
Tall-03 = Caballero
Tall-04 = Plus
Tall-05 = Unica
Tall-06 = Accesorio
```

Es decir: lo que el sistema llama **"Talla" no es una talla de prenda (S/M/L/XL)** — es una **categoría de corte/consumo** que determina cuánta tela se necesita. Confirma directamente la hipótesis del usuario en la sección 6 del pedido, con un matiz importante que la hipótesis no anticipaba: la misma columna también absorbe **"Unica"** (para productos sin variación de tamaño) y **"Accesorio"** (para productos que no son prendas en absoluto — bandanas, tulas). Es decir, "Talla" en este negocio es en realidad **"categoría de consumo/corte"**, sobrecargada con un cuarto uso (marcador de "no aplica tamaño").

**Y existe una segunda dimensión de talla, completamente distinta, para el precio (no el costo):** el texto literal en `PresupuestoPlantilla!B24`:

> *"EL PRECIO A PARTIR DE LA TALLA XL, SE LE SUMA $3 AL PRECIO POR UNIDAD (EJEMPLO: TALLA L $15 TALLA XL $18, TALLA 2XL $21)."*

Esta es la talla de venta al público (S/M/L/XL/2XL), y es **[REGLA CONFIRMADA]**: desde XL se suma un recargo fijo de USD 3 al precio de venta (no al costo). Es una regla de **precio**, no de **consumo**. El negocio opera con **dos escalas de "talla" distintas y no relacionadas explícitamente en las fórmulas**: una que determina consumo de tela (Niño/Dama/Caballero/Plus/Unica/Accesorio) y otra que determina recargo de precio (S/M/L/XL/2XL). Ninguno de los tres archivos modela la relación entre ambas — probablemente porque en la práctica el operador simplemente sabe cuál aplica a cada pedido.

### 4.1 La tabla maestra de consumo tiene huecos reales sin resolver

`Propuesta.xlsm!Tabla4 (3)` (= `Consumo_Tela_Producto`, la fuente que consume `obtenerconsumotela()`), 36 filas, producto+talla→consumo en m². De ellas, **7 filas (~19%) tienen dos o tres valores en la misma celda**, sin resolver cuál es el vigente:

| Producto | Talla | Valor en la celda |
|---|---|---|
| P-03 | Dama | `0,9 Y 1,1` |
| P-03 | Caballero | `1,5 Y 2` |
| P-03 | Niño | `0,8 Y 0,9, 0,5/1,5` |
| P-04 | Caballero | `1,1-1,6` |
| P-08 | Caballero | `2,2 Y 2,5` |
| P-15 | Dama | `0,6/2; 0,8` |
| P-15 | Niño | `0,5, 0,6/2` |

**[DICE]**. Rastreado hasta su origen en `extraccion manual!X3`: el precio de la tela "muselina" también aparece como `"1,5 Y 1,8"` — dos precios, nunca resueltos a uno solo. Esto sugiere (**[HIPÓTESIS]**) que la ambigüedad no es solo de consumo: pudo originarse en que la misma prenda se cotizó en momentos distintos con anchos de rollo o proveedores de tela distintos, y ambos números quedaron pegados en la misma celda en vez de crear dos filas.

**Consecuencia verificable en el motor**: `obtenerconsumotela()` hace `XLOOKUP(...) → Cantidad_Tela_m²`. Si esa celda contiene el texto `"0,9 Y 1,1"` en vez de un número, la multiplicación posterior (`consumo × precio_final`) produce un error de tipo en Excel. **Esto no es una hipótesis — es una consecuencia mecánica verificable de la fórmula tal como está escrita contra el dato tal como está guardado.** Cotizar P-03 en talla Dama, Caballero o Niño con el motor actual está roto para esas combinaciones específicas. Clasificación: **A (error de implementación) causado por C (regla de negocio implícita nunca resuelta)** — la ambigüedad del dato viene de una realidad de negocio no capturada (dos telas/anchos posibles), pero su síntoma es un error de fórmula real.

### 4.2 La llave de producto no es consistente

`Tabla4 (3)!A2 = "KASAKAS "` y `!A36 = "Tula"` — nombres de producto en texto libre, no códigos (`P-XX`) como el resto de las 35 filas. **[DICE]**. `obtenerconsumotela()` filtra por `Codigo_Producto = prod`, así que estas dos filas nunca hacen match si se les pasa un código real — quedan huérfanas del motor salvo que alguien busque exactamente ese texto. Clasificación: **B (dato inconsistente)** — probablemente entradas agregadas en un momento distinto, sin pasar por el mismo proceso de normalización que las demás.

---

## 5. Modelo de costos

Componentes de costo confirmados, con su unidad de referencia:

| Componente | Cómo se calcula | Moneda base | Evidencia |
|---|---|---|---|
| Tela | consumo_m² × precio_material × (tasa si USD) | USD (convertido) | `obtenerCostoTela`, `fRANELA!B4` |
| Vinil | área de anidado (ancho útil rollo / diseño, por filas) × precio material | USD (convertido) | `obtenerCostoVinil` + `ObtenerConsumoVinil` |
| Sublimado | consumo por tamaño de estampado o por tela (mitad si "delantero", total si "total") × precio material | USD (convertido) | `obtenerCostoSublimado`, coincide con nombres "Sublimado Delantero"/"Sublimado Total" en `Produccion` |
| DTF | igual patrón que sublimado, con su propia función | USD (convertido) | `ObtenerCostoDTF` |
| Actividades (corte, costura, planchado...) | tarifa fija por producto + tramo de volumen, con `P-12` como tarifa genérica de respaldo si el producto no tiene tarifa propia | COP casi siempre | `obtenerCostoActividad`, `fact_costo_actividad_producto` |
| Consumibles por rendimiento (láminas, rollos) | precio de la lámina/rollo ÷ rendimiento en unidades | COP | `fRANELA!B89=13000/15` y similares |

**[REGLA CONFIRMADA]** — Materiales tiene **doble precio** (Detal / Mayor): `obtenerCostoMaterial` usa Mayor cuando el tramo de volumen no es "Detal" (VC001) y Mayor>0, si no cae a Detal. Coincide con `propuesta-COTIZADOR.xlsx!Materiales` (`Costo_Unitario_Detal`, `Costo_Unitario_Mayor`).

**[REGLA CONFIRMADA]** — Margen inversamente proporcional al volumen: `PresupuestoPlantilla!S17:U17 = 0.99 / 0.55 / 0.4` (99%, 55%, 40%), aplicados como `Subtotal × (1+margen)` en `S18/T18/U18`. Pedidos pequeños llevan más margen, pedidos grandes menos.

**[REGLA PROBABLE, no confirmada al 100%]** — Los tramos de volumen **no usan el mismo corte en todas partes**: el motor central corta en 12 y 1000 unidades (`K7`: <12→VC001, <1000→VC002, resto→VC003), mientras que `fRANELA...xlsx` usa columnas "+5 DOC / 1 A 5 DOC / DETAL" (cortes en docenas, no en unidades sueltas: 60 y algo menos). No hay evidencia de que ambos esquemas sean el mismo tramo expresado en unidades distintas — parecen ser dos definiciones de "volumen" que coexistieron en momentos distintos del negocio. **Pendiente de confirmar con el dueño del negocio** (ver sección 13).

---

## 6. Reglas de negocio descubiertas

### Confirmadas (evidencia directa en fórmula o texto)
1. Costo de tela = consumo(m) × precio_USD/m del material × tasa de cambio, cuando el material está en USD.
2. El costo total del pedido se calcula en COP línea por línea y se normaliza a USD dividiendo entre la tasa (`Total_COP / tasa = costo_unitario_USD`).
3. Desde talla XL se suma un recargo fijo de $3 USD al precio de venta por unidad (texto literal en el original).
4. Se requiere abono del 60% del total para iniciar producción (texto literal).
5. El margen de venta baja según sube el volumen del pedido (99%→55%→40%).
6. Los materiales tienen precio Detal y precio Mayor; se usa Mayor solo si el tramo de volumen no es Detal y el Mayor está definido.
7. Si un producto no tiene tarifa propia de actividad, se usa la tarifa del producto genérico `P-12`.
8. "Talla" en el motor de consumo de tela es categoría Niño/Dama/Caballero/Plus/Unica/Accesorio, no una talla de prenda.
9. El consumo de vinil depende del tamaño del diseño y de la cantidad pedida, vía cálculo de anidado sobre un ancho útil de rollo fijo (48 cm).
10. Existe un mecanismo de override por línea de cotización que no toca el maestro (columna "Unidad Manual" — ver sección 8).
11. **[Validado por la dueña, 2026-08-17]** El ancho de tela útil NO es un valor fijo (~1,50 m) igual para todas las telas — varía según el tipo de tela y el patrón de corte de cada prenda. El consumo real además se reduce por **merma de corte y costura** ("se reducen al cortar, al coser y todo eso"): la tela sobrante o perdida en esos procesos ya está descontada en el número final de consumo, no es un factor aparte que el sistema calcule.
12. **[Validado]** La categoría de consumo (Niño/Dama/Caballero/Plus/Unica/Accesorio) **no se deriva automáticamente de la talla comercial** — la dueña la asigna por criterio propio, según cuánta tela necesita realmente esa prenda en ese lote: "si ella cose tantas S, ella las puede meter como Niño porque ella así tiene programada su cantidad de tela". Hay una correspondencia aproximada y no estricta (S/M ≈ Dama; L/XL ≈ Caballero), pero la decisión final es manual, por lote. **Consecuencia para el modelo**: no construir una tabla fija TallaComercial→CategoriaConsumo; el operador debe poder elegir la categoría de consumo directamente, como ya hace el cotizador actual.
13. **[Validado]** El recargo de +$3 desde XL existe porque XL (y "Plus", que la dueña cree equivalente a XL en este punto) consume más tela que las tallas menores — es un recargo de precio justificado por mayor consumo real, no una regla arbitraria.
14. **[Validado]** El abono del 60% es protección contra pérdida: una vez iniciada la producción (tela cortada, procesos empezados), una cancelación del cliente deja a la dueña con material y trabajo ya invertido sin poder recuperarlo.
15. **[Validado, con matiz importante]** El margen decreciente por volumen (99%/55%/40%) **no es únicamente política de precio — es recuperación del desperdicio real de material**: al cortar una sola pieza de vinil (o tela), se pierde la tira completa de 15/20/30 cm sobrante igual que si se hubieran cortado varias piezas de la misma tira. A menor cantidad, mayor desperdicio proporcional por unidad, y el margen más alto compensa eso, no es solo "ganar más en pedidos chicos".
16. **[Validado]** El ancho útil de vinil (48 cm) viene de un rollo físico de 50 cm menos el margen que se pierde en los rodillos de la cortadora de plotter.
17. **[Validado]** Los insumos secundarios (hilo, marquilla, bolsa, etiqueta) efectivamente no varían por producto — pero **sí varían por color/ítem específico de inventario**, y ese es el motivo real, cotidiano, del mecanismo de override (ver punto 18).
18. **[Validado — hallazgo nuevo, no anticipado por el pedido original]** **El sistema no modela color como dimensión**, ni en el original ni en el refactor. La dueña lo señaló como una carencia real: accesorios y avíos (cierres, cuellos, etc.) cambian de costo según el color específico que haya que comprar. Ejemplo textual: un pedido de camisa morada con cierre morado que no había en inventario obligó a comprar ese color por separado, a un costo más alto que el habitual por comprarse una sola unidad. **Esta es la causa principal, real y frecuente del override** ("pueden haber muchos casos") — no es un caso raro de excepción, es el mecanismo cotidiano para absorber variación de costo por color/disponibilidad de inventario específico.
19. **[Validado]** El rendimiento de consumibles de sublimación (precio de lámina ÷ diseños que rinde) es un valor estándar y fijo; cuando sube el precio del material, se actualiza una sola vez en la tabla de materiales, no se recalcula el rendimiento por receta.

### Probables (evidencia fuerte, no verificación cruzada completa)
20. Los tramos de volumen del motor central (12 / 1000 unidades) y los de las hojas manuales (docenas) siguen sin confirmarse como el mismo criterio o dos criterios distintos — la dueña no llegó a aclarar este punto específico en su respuesta.

### Hipótesis (necesitan más verificación — ya no completamente abiertas)
21. **[Hipótesis revisada, con nueva evidencia parcial]** Las celdas con dos valores de consumo de tela ("0,9 Y 1,1" en `Tabla4 (3)`) — la dueña no las había visto todavía al momento de responder ("tendrías que ver qué tienen esas dos filas... ya voy a mirar el archivo"). Una revisión dirigida de `HOJA DE CALCULO DE PRECIOS AL MAYOR.xlsx!BUZO` (sheet correspondiente a P-03) muestra el mismo producto+categoría con **bloques separados por tipo de tela** (Muselina: `0.9*1.8`; Muselina variante "réplica": `1.1*1.6`; Licra: `0.9*4`; Micro: `1*2.3`...) — es decir, cada valor de la celda ambigua probablemente corresponde a un tipo de tela distinto que, en algún momento de consolidación anterior, se combinó en una sola celda en vez de mantenerse como filas separadas por tela. Sigue sin ser una confirmación completa (no se verificaron las 7 filas una por una), pero cambia el diagnóstico: no parece ser una contradicción entre dos mediciones del mismo consumo, sino **pérdida de la dimensión "tipo de tela" al consolidar la tabla**.
22. "Sublimado Delantero" consume la mitad de tela que "Sublimado Total" porque solo estampa el frente de la prenda (inferido del nombre y de que `ObtenerConsumoDTF` divide entre 2 cuando el tamaño es "TAM-08") — sin verificar todavía.

---

## 7. Recetas — ¿hay evidencia suficiente?

**Sí, y ya está parcialmente implementada como código, no solo como idea.** `obtenerCostoGlobal` es literalmente un despachador de "receta de técnica": según el código de actividad, llama a una función distinta (`obtenerCostoVinil`, `obtenerCostoSublimado`, `ObtenerCostoDTF`, `obtenerCostoActividad` genérico, `obtenerCostoMaterial`), cada una con su propia lógica de consumo. Esto **es** la separación receta-de-producto + receta-de-técnica que planteaba la hipótesis del pedido, ya expresada como código real, no solo como intención.

Lo que **falta** para que sea una receta completa y declarativa (hoy vive como lógica de programa, no como datos):
- No hay una tabla `Receta_Producto` que enumere, para cada producto, qué componentes (tela, corte, costura, etiqueta, hilo, bolsa, marquilla, gastos operativos) le aplican y en qué cantidad — hoy esa lista está hardcodeada como filas fijas dentro de `PresupuestoPlantilla` (K11:K27), igual para todo producto.
- Los insumos secundarios (hilo, marquilla, bolsa, etiqueta) **no varían por producto en las hojas manuales** — son casi siempre el mismo valor (250, 750, 100, 200) repetido en cada bloque de `fRANELA...xlsx`, lo que sugiere que sí son genéricos y no necesitan una receta por producto, solo una tarifa por defecto.
- La receta de técnica (vinil/sublimado/DTF) sí depende del tamaño elegido (`Dim_Tecnica_Tamano`), y ese componente **sí** necesita ser una tabla, porque los tamaños con sus cm de ancho/alto varían por técnica.

**[PROPONGO]** — no implementar todavía, pero anotar como dirección: separar "receta base del producto" (tela + operaciones fijas de corte/costura/insumos genéricos) de "receta de técnica" (el tamaño y material de la decoración elegida), exactamente como ya lo hace el código LAMBDA, pero moviendo esa lógica de fórmula a datos configurables.

---

## 8. Overrides — SÍ existen en el original, y NO se migraron al refactor

**[REGLA CONFIRMADA, con mecanismo preciso]**. En `Propuesta.xlsm!PresupuestoPlantilla`, la tabla `Costro_Producto` tiene columnas `Unidad` (costo calculado) y `Unidad Manual` (override), y la fórmula de cada línea es:

```
O11 = IFERROR(IF(Costro_Producto[Unidad Manual]<>"", Costro_Producto[Unidad Manual], Costro_Producto[Unidad]), )
```

Es decir: si el operador escribe un valor en "Unidad Manual" para esa línea, se usa ese valor; si la deja vacía, se usa el costo calculado por el motor. **Esto es exactamente el mecanismo que se pidió preservar**: "la tela cuesta $2 en el catálogo pero esta vez costó $3" se resuelve escribiendo 3 en "Unidad Manual" de esa línea, sin tocar `HechosInventario` (el maestro de materiales).

**Verificado que no sobrevivió al refactor**: revisé `propuesta-COTIZADOR.xlsx!Presupuesto` y el módulo `M01_GestionPresupuesto.bas` (VBA que gestiona esa hoja). El precio de cada línea agregada a `T_Items` viene siempre de `L17` (calculado), sin ningún campo equivalente a "Unidad Manual". El botón "Cargar para editar" permite reabrir una línea y volver a calcularla, pero no anular puntualmente el resultado del motor manteniendo el resto de la lógica. **Esta es una pérdida real de una capacidad de negocio que el original sí tenía y documentaba con un ejemplo funcionando** — no una simplificación cosmética.

**[PROPONGO]** para el modelo futuro: mantener la separación catálogo/cotización explícita en el modelo de datos, no solo como columna extra. Cada línea de cotización tiene un costo_calculado (derivado de las reglas del catálogo en el momento de cotizar) y un costo_override opcional; el override nunca se escribe hacia el catálogo. Lo mismo aplica a precios de técnica (el ejemplo del usuario: sublimación normalmente $20.000, para un amigo $10.000 solo en esa cotización).

---

## 9. Inconsistencias del original, clasificadas

| # | Hallazgo | Categoría | Evidencia |
|---|---|---|---|
| 1 | 7 de 36 filas de `Consumo_Tela_Producto` tienen dos/tres valores en la misma celda | **A + C** (error de implementación causado por regla implícita no resuelta) | `Tabla4 (3)!P8, Q8, R8, Q9, Q13, P26...` |
| 2 | Dos filas de esa misma tabla usan nombre de producto en vez de código | **B** (dato inconsistente) | `Tabla4 (3)!A2="KASAKAS ", A36="Tula"` |
| 3 | Tasa de cambio referenciada (`E2`/`E140`/`K1`) en unas fórmulas, hardcodeada (3700/3800/3900) en otras, en el mismo archivo | **B + D** (dato inconsistente / posible decisión manual puntual al momento de cotizar) | Ver sección 3.2 |
| 4 | Dos esquemas de tramo de volumen distintos (12/1000 unidades vs. docenas) | **C** (regla implícita, no reconciliada) | `PresupuestoPlantilla!K7` vs. `fRANELA...` columnas C-E/H-J |
| 5 | `CostoMaterialDecorativo` es una función LAMBDA que siempre devuelve el texto `"Prueba2"` | **Código inacabado**, no encaja en las 6 categorías — lo marco aparte | Definición de nombre en `Propuesta.xlsm` |
| 6 | El nombre definido `ObtenerConsumoDTF` tiene su cuerpo LAMBDA duplicado y concatenado con `=` en la definición cruda | **F** (ambigüedad técnica — no está claro si es un artefacto de guardado o un error real de Excel) | `NombresDefinidos`, fila 124 |
| 7 | Precio de "muselina" registrado como dos valores ("1,5 Y 1,8") desde el origen (`extraccion manual`) | **E o C** — probablemente dato histórico de dos cotizaciones distintas nunca reconciliado | `extraccion manual!X3` |
| 8 | Cálculos sueltos sin etiqueta clara al margen de un bloque (`K245:K247` en `fRANELA...`: resta, división por tasa, multiplicación por 30%) | **D** (decisión manual puntual para una cotización específica) | `fRANELA...!H246:K247` |

No encontré evidencia de que ninguna de estas inconsistencias sea simplemente "el Excel está mal hecho" en el sentido de estar vacío de intención — cada una tiene un origen de negocio identificable, aunque algunas (1, 3, 4, 7) representan realidades que el archivo nunca terminó de capturar limpiamente.

---

## 10. Comparación con el cotizador refactorizado — qué ganó, qué perdió

| Aspecto | Original (`Propuesta.xlsm`) | Refactor (`propuesta-COTIZADOR.xlsx` + VBA) | Veredicto |
|---|---|---|---|
| Estructura relacional | Mezclada: motor LAMBDA limpio conviviendo con hojas de trabajo manual desordenadas | Tablas separadas por dominio (Productos/Materiales/Producción/TasasCambio) | **Mejora real** |
| Tasa de cambio | Centralizada en el motor LAMBDA (parámetro `tasa_dolar`), pero hardcodeada en hojas manuales paralelas | Centralizada en hoja `TasasCambio`, referenciada consistentemente | **Mejora real** |
| Consumo de tela por producto+talla | Tabla `Consumo_Tela_Producto`, con huecos (sección 4.1) | Tabla `Produccion!AK:AO` con la misma estructura, mismos datos, **mismos huecos no resueltos** (columna `Cantidad_Tela_Formula` existe como encabezado pero está vacía en todas las filas revisadas) | **Heredó el problema, no lo resolvió** |
| Consumo de vinil por anidado geométrico | `ObtenerConsumoVinil` — cálculo real de piezas por fila y filas necesarias | No encontré una función equivalente en `propuesta-COTIZADOR.xlsx` ni en el VBA | **Pérdida, sin confirmar si fue intencional** |
| Precio Detal/Mayor de materiales | Dos columnas | Dos columnas (`Costo_Unitario_Detal`/`Mayor`) | **Se conservó** |
| Tramos de volumen para técnica/actividad | 3 tramos (Detal/+12/+1000) | 3 tramos (`VC001/VC002/VC003` = Detal/Docena+12/Mayor+1000) | **Se conservó**, con nombres más claros |
| Override por línea sin tocar el maestro | Columna "Unidad Manual", funcional | **No existe** en el flujo de `M01_GestionPresupuesto.bas` | **Pérdida real y documentada** — ver sección 8 |
| Recargo por talla (+$3 desde XL) | Texto explícito en condiciones | No encontré ninguna fórmula ni tabla que lo reproduzca | **Pérdida, o pendiente de implementar** |
| P-12 como tarifa genérica de respaldo | Implementado en `obtenerCostoActividad` | Existe el producto `P-12 "Genérico (todos)"` en el catálogo, pero no verifiqué si `fact_costo_actividad_producto` en `Produccion` realmente cae a él cuando falta una tarifa específica | **Sin confirmar — riesgo** |

**Autocrítica pedida explícitamente**: el trabajo de refactorización (hecho en la sesión anterior de este mismo proyecto) mejoró genuinamente la estructura de datos, pero **simplificó de más en al menos dos puntos con impacto de negocio real**: el override por línea (una capacidad que el dueño del negocio usa activamente, según el propio texto del original) y el cálculo de anidado de vinil (que evita sobreestimar o subestimar el consumo real del material). Ninguna de las dos pérdidas está documentada en `DECISIONS.md` ni en `DEFERRALS.md` — no fueron decisiones conscientes registradas, lo que sugiere que se perdieron por no haberse detectado, no por haberse descartado a propósito.

---

## 11. Matriz de paridad

| Funcionalidad del original | ¿Existe en refactorizado? | ¿Se comporta igual? | ¿Se perdió información? | ¿Debe conservarse? | Observaciones |
|---|---|---|---|---|---|
| Consumo de tela por producto+talla | Sí | Sí | No (mismos huecos) | Sí | Resolver los 7 casos ambiguos antes de dar por bueno |
| Consumo de vinil por anidado geométrico | No encontrado | — | Sí | Sí | Reconstruir como función/regla explícita |
| Consumo de sublimado/DTF por tamaño o por tela | Parcial (tabla de tamaños existe) | Sin confirmar | Posible | Sí | Falta ver si la lógica "delantero=mitad" se reprodujo |
| Ancho de tela ≈1,50 m | No aparece como constante explícita en ningún archivo | — | — | A confirmar | Ver pregunta crítica 1 |
| Precios en USD y COP mezclados | Sí | Sí, mejor resuelto | No | Sí | Refactor con tasa centralizada es mejor |
| Conversión de moneda | Sí (parcial en original) | Sí (consistente) | No | Sí | Mejora real del refactor |
| Tramos de volumen (Detal/Docena+12/Mayor+1000) | Sí | Sí | No | Sí | — |
| Precio Detal/Mayor por material | Sí | Sí | No | Sí | — |
| Margen inverso al volumen (99/55/40%) | Sí | No encontrado en el refactor | Sí | Sí | Verificar si se movió a otro lugar no revisado |
| Recargo +$3 desde talla XL | Sí (texto) | No encontrado | Sí | A confirmar | ¿Sigue vigente o es histórico? |
| Override manual por línea sin tocar catálogo | Sí, funcional | No | Sí | **Sí, crítico** | Ver sección 8 |
| Tarifa genérica P-12 de respaldo | Sí | Catálogo existe, motor sin confirmar | Posible | Sí | — |
| Abono 60% para iniciar | Sí (texto) | Estructura de Abono existe en `Presupuesto`/`Comandas` | No en estructura, sin confirmar si el 60% se aplicó como regla o queda a criterio manual | A confirmar | — |
| Costeo de consumibles por rendimiento (lámina÷diseños) | Solo en hoja manual (archivo 3) | No encontrado en ninguno de los otros dos | Sí, si es una regla real y no un cálculo de una sola vez | A confirmar | Ver pregunta crítica 4 |

---

## 12. Diccionario del negocio

| Término | Qué significa realmente | Cómo aparece en el Excel | Dónde se usa | Ambigüedades | Cómo modelarlo después |
|---|---|---|---|---|---|
| **Producto** | Tipo de prenda o artículo del catálogo (Franelas, Bandanas, Buzos...) | `Dim_Producto[Codigo_Producto, Nombre, Categoría, Subcategoría]` | Todo el motor | Ninguna relevante | Tabla maestra simple |
| **Talla (de consumo)** | Categoría de corte que determina cuánta tela se usa: Niño/Dama/Caballero/Plus/Unica/Accesorio | `Dim_Talla[Id_talla, Talla]` | `Consumo_Tela_Producto`, motor de costeo | Se confunde con la talla de venta (S/M/L/XL) | Renombrar conceptualmente a "categoría de consumo" o similar; separar de talla comercial |
| **Talla (de venta/precio)** | Talla comercial de la prenda (S, M, L, XL, 2XL...) | Solo texto libre en `PresupuestoPlantilla!B24`, sin tabla propia encontrada | Recargo de precio (+$3 desde XL) | No hay tabla ni relación formal con la Talla de consumo | **PENDIENTE DE CONFIRMACIÓN** — necesita su propia dimensión |
| **Tela** | Material textil consumido para fabricar la prenda | `HechosInventario` filtrado por nombre que contiene "Tela"; consumo en `Consumo_Tela_Producto[Cantidad_Tela_m²]` | `obtenerCostoTela` | Precio a veces registrado con dos valores | Tabla de materiales con flag `EsTela`, como ya hace el refactor |
| **Ancho de tela** | Ancho físico del rollo de tela (hipótesis: ~1,50 m) | No encontré una constante explícita con ese valor en ningún archivo | — | **No confirmado en los archivos** — el usuario lo dio como hipótesis de partida, pero no aparece como número reconocible en ningún lado | **PENDIENTE DE CONFIRMACIÓN** con el dueño del negocio |
| **Consumo** | Cantidad de tela (en metros o m²) necesaria para una unidad de producto+talla | `Cantidad_Tela_m²` | `obtenerconsumotela` | 7 filas con doble valor sin resolver | Fact table producto×talla→consumo, resolviendo antes la ambigüedad |
| **Técnica** | Método de decoración: Vinil, Sublimado, DTF, Bordado | `dim_Tecnica[Id_Tecnica, técnica]` | Dispatcher `obtenerCostoGlobal` | Ninguna relevante | Catálogo simple + función de costeo propia por técnica |
| **Tamaño (de técnica)** | Dimensión física del área decorada (15cm, Carta, Oficio, 20x20...) | `Dim_Tecnica_Tamano[Id_Tamano, Ancho_cm, Alto_cm]` | Consumo de vinil/sublimado/DTF | Ninguna relevante | Tabla propia por técnica |
| **Receta** | No existe como tabla — vive como lógica de programa en el dispatcher LAMBDA | — | — | — | Ver sección 7 |
| **Insumo** | Material secundario no textil: hilo, marquilla, bolsa, etiqueta | `HechosInventario`, mayormente valores fijos independientes del producto | Líneas fijas del cotizador | Ninguna relevante | Catálogo con tarifa por defecto, override si un producto lo necesita distinto |
| **Costo** | Valor interno de producir una unidad, calculado siempre primero en COP | `Costro_Producto[Sub Total]`, `Total` | Todo el motor | Ninguna | — |
| **Precio** | Costo + margen, mostrado en la moneda elegida por el cliente | `PresupuestoPlantilla!L17`, tabla `S18:U19` | Presentación al cliente | Ninguna | — |
| **Precio maestro** | El costo/precio de catálogo (material, técnica) sin ajustar a una cotización particular | `HechosInventario[Costo_Unitario_Detal/Mayor]` | Motor de costeo | — | Nunca se modifica desde una cotización |
| **Override** | Valor que reemplaza el costo calculado solo para una línea de una cotización específica | Columna `Unidad Manual` en `Costro_Producto` | `PresupuestoPlantilla!O11` | **No existe en el refactor** | Ver sección 8 |
| **Moneda** | USD (interna/costeo) y COP/VES (según el cliente) | `HechosInventario[Moneda]`, `Presupuesto!C7` | Todo | — | Catálogo de monedas + tasas, ya resuelto razonablemente en el refactor |
| **Tipo de cambio / tasa** | Valor de conversión USD↔COP/VES | `K1` (original), `E2`/`E140` (archivo 3, sin centralizar), `TasasCambio` (refactor) | Todo cálculo en USD | Inconsistente en el original, resuelto en el refactor | Tabla de tasas con fecha de vigencia (no solo un valor fijo) |
| **Cantidad** | Unidades pedidas en una línea de cotización | `PresupuestoPlantilla!K6`, `L9` | Determina tramo de volumen | — | — |
| **Tramo de volumen** | Rango de cantidad que determina tarifa/margen (Detal, Docena+12, Mayor+1000) | `VC001/VC002/VC003` | Costeo de actividades y margen | Dos definiciones distintas conviven (unidades vs. docenas) | **PENDIENTE DE CONFIRMACIÓN** |

---

## 13. Preguntas críticas para la dueña/dueño del negocio

**Estado tras la respuesta del 2026-08-17**: de las 10 preguntas originales, **8 quedaron resueltas** (1, 3, 4, 6, 7, 8, 9 confirmadas; 4 con matiz importante en el caso de los tramos de margen). Quedan dos abiertas.

### Resueltas por la auditoría completa (sección 16)
1. **(antes #2, "los 7 casos ambiguos")** — **Resuelto.** No son valores alternativos: son residuos de consolidar bloques de distintos tipos de tela (o los dos factores de una multiplicación) en una sola celda. Ver sección 16.1.
2. **(antes #5, "tramos de volumen")** — **Parcialmente resuelto, con hallazgo más amplio de lo esperado**: no son dos esquemas, son **al menos cuatro** nomenclaturas distintas para el mismo concepto de 3 tramos (`VC001/VC002/VC003` con corte en 12/1000 unidades; `+5 DOC/1 A 5 DOC/DETAL`; `Unidad/Docena/Mayor`; `1/6 a 12/12 o más`), y los umbrales exactos varían por hoja. No hay un estándar único documentado en ningún archivo — cada producto/vendedor parece haber definido el suyo con el tiempo. Sigue pendiente decidir con la dueña **cuál es el corte correcto para el sistema nuevo**, pero ya no es una duda de "cuál de dos" sino "cuál definir de cero".

### Resueltas — respuesta registrada como regla confirmada
- ~~¿Ancho de tela fijo?~~ → No, varía por tela; el consumo ya descuenta merma de corte/costura. Ver regla #11.
- ~~¿Relación entre las dos tallas?~~ → No hay tabla fija; la dueña asigna la categoría de consumo por criterio propio, por lote. Ver regla #12.
- ~~¿Vigencia de recargo XL y abono 60%?~~ → Ambos vigentes y con justificación de negocio real (consumo de tela y protección contra cancelación). Ver reglas #13-14.
- ~~¿Frecuencia del override?~~ → Frecuente, y su causa principal es variación de costo por color/inventario específico — un hallazgo nuevo, no anticipado. Ver reglas #17-18.
- ~~¿Vigencia del ancho útil de vinil 48cm?~~ → Confirmado: rollo de 50cm menos margen de los rodillos del plotter. Ver regla #16.
- ~~¿Insumos secundarios varían por producto?~~ → No por producto, sí por color/inventario. Ver regla #17.
- ~~¿Rendimiento de sublimación fijo o variable?~~ → Fijo, se actualiza una sola vez en la tabla de materiales si sube el precio. Ver regla #19.

### Nueva pregunta crítica (surgida de esta conversación, no estaba en la lista original)
11. **¿Se necesita una auditoría completa de `HOJA DE CALCULO DE PRECIOS AL MAYOR.xlsx` (51 hojas, una por línea de producto) como parte de esta Fase 0, o el alcance del análisis se cierra con lo ya cubierto (Franelas + motor de costeo + reglas validadas) y esa hoja se usa solo como referencia puntual cuando haga falta verificar un producto específico?** Es una decisión de alcance/tiempo, no de negocio — la escala de ese archivo es varias veces la de los tres archivos originales.

### Opcional, sigue abierta
12. Los cálculos sueltos sin etiquetar en el margen de `fRANELA...xlsx` (fila 245-247) — ¿corresponden a una cotización real que vale la pena rescatar como caso de prueba, o son un cálculo descartable?

---

## 14. Modelo conceptual propuesto (sin código, sin SQL)

**[PROPONGO]** — a partir de lo confirmado en este análisis, no de una plantilla genérica:

```
CATÁLOGOS
  Producto            (código, nombre, categoría)
  CategoriaConsumo     (antes llamada "Talla" en el original — Niño/Dama/Caballero/Plus/Unica/Accesorio;
                        se asigna por criterio del operador en cada lote, no por una tabla fija)
  TallaComercial        (S/M/L/XL/2XL — dimensión separada, ligada a CategoriaConsumo solo como referencia aproximada)
  Material              (tela e insumos, con precio Detal/Mayor, moneda, flag EsTela)
  Color                 (validado como dimensión real y ausente en ambos archivos: cambia el costo real de
                        avíos/accesorios cuando hay que comprar un color específico fuera de inventario habitual)
  Tecnica               (Vinil/Sublimado/DTF/Bordado)
  TamanoTecnica          (por técnica: ancho/alto del área decorada)
  Moneda + TasaCambio    (con fecha de vigencia, no un valor fijo sin historial)
  TramoVolumen           (Detal/Mayor — con una sola definición, resolviendo la pregunta crítica 5)

RECETAS
  ConsumoTela            (Producto × CategoriaConsumo → metros de tela — resolviendo antes los 7 casos ambiguos)
  RecetaBaseProducto      (operaciones fijas: corte, costura, insumos genéricos con su tarifa por defecto)
  RecetaTecnica            (por técnica: cómo se calcula su consumo — área/anidado para vinil, tamaño o tela para sublimado/DTF, rendimiento por lámina para consumibles)

COTIZACIÓN
  Cotizacion (cabecera: cliente, fecha, moneda de presentación, estado)
  Cotizacion_Linea (producto, categoría de consumo, técnica, tamaño, material, cantidad,
                    costo_calculado, costo_override [opcional, nunca escribe al catálogo],
                    margen_aplicado según tramo de volumen)
  Comanda / Abono / Saldo (igual que el original, generalizado)
```

La diferencia clave frente al modelo que ya se había decidido en la sesión anterior (D-001, "atributos genéricos configurables") es que **este análisis no encontró evidencia suficiente en los tres archivos para justificar generalizar ya a un modelo de atributos abstractos** — lo que sí hay evidencia sólida es de un dominio textil concreto con tres patrones de consumo distintos (tela por m², técnica por área/anidado, consumible por rendimiento) que conviene modelar bien primero en su forma concreta. Generalizar a "atributos configurables" antes de tener esto sólido es exactamente el riesgo que la sesión anterior ya corrió con D-001, apoyada en un archivo (`RECETAS_NORMALIZADAS.xlsx`) que ya no está disponible para verificar.

---

## 15. Riesgos de implementación (qué se rompe si se construye ya)

1. **Construir sobre `Consumo_Tela_Producto` tal cual, sin resolver los 7 casos ambiguos, hereda un motor que falla silenciosamente o con error para esas combinaciones específicas** — ya le pasó al refactor anterior, que copió el mismo hueco sin corregirlo.
2. **Si se retoma el modelo de "atributos genéricos" (D-001) sin antes fijar el modelo concreto de tela/técnica/consumo, se corre el riesgo de abstraer sobre una base todavía inestable** — la generalización es prematura mientras sigan sin resolver las preguntas críticas 1-5.
3. **Omitir el override por línea (como ya pasó en el refactor) elimina una capacidad que el negocio usa y documenta explícitamente** — si se construye el sistema nuevo sin este campo desde el principio, agregarlo después obliga a tocar todo el flujo de cotización otra vez.
4. **Confundir las dos escalas de "talla" (consumo vs. comercial) en un solo campo** produciría un sistema que no puede aplicar el recargo de XL correctamente ni calcular consumo correctamente al mismo tiempo.
5. **Perder el cálculo de anidado de vinil** (reemplazarlo por área simple × cantidad) subestima o sobreestima el material real en pedidos con muchas piezas pequeñas — impacto directo en costo real vs. costeado.

---

## 16. Auditoría completa de `HOJA DE CALCULO DE PRECIOS AL MAYOR.xlsx` (51/51 hojas, 2026-08-17)

Revisadas las 51 hojas (una por línea de producto: BUZO, SUETER, FRANELILLAS, SHORT, LICRAS, UNI FUTBOL, CAMISAS, CV TASLAN, CV MEMORY, CHALECOS, FRA ALGODON, HODDI MONO/TERRY, PANTALON MONO, KSAKAS, F. SUBLIM, FULL SUBLI, PIJAMAS, CANILLERAS, FUNDACANILLERA, TULA, CHEMIS, BRAGA, BEISBOL, VELAS, BOLSAS POP, CAMISAS, FILIPINA, VINILES, MONO QUIRURGICO, TUTU, DELANTAL, FRANELILLAS, MANGAS, GORRA, TOP DAMA, COJINES, VESTIDO PROMOTORA, WALL/WALL MEDIA, LANYARD, franelas cucuta, TAPASOL, FALDA SHORT LICRA, CROPTOP LICRA, MANUBRIO, ESTANDARTE, MOUSE, PAÑOLETA, PROMO, P COSTURA, y FRA MUSE MICR — esta última es la misma hoja que el archivo 3 original). Confirma el mismo patrón de cálculo en el 100% de las hojas (`consumo × precio × tasa`, 3 tramos de volumen, override manual puntual) y agrega hallazgos que no eran visibles con solo tres archivos:

### 17.1 Resuelto: el origen de los 7 casos de consumo "ambiguo" en `Tabla4 (3)`

Confirmado con evidencia directa, no solo hipótesis. Los valores dobles/triples **no son dos mediciones alternativas del mismo consumo** — son residuos de una consolidación que perdió una dimensión real: **el tipo/marca específica de tela**. Dos mecanismos distintos, ambos verificados:

- **Los dos números eran los dos factores de una sola fórmula de multiplicación**, no dos consumos. Ejemplo exacto: `Tabla4 (3)` registra P-08/Caballero = `"2,2 Y 2,5"`; en `CV TASLAN!B25`, la fórmula real es `=2.2*2.5*F4` — 2.2 y 2.5 son ancho y largo de la MISMA prenda, no dos alternativas.
- **Los dos números venían de dos bloques de tela distintos para el mismo producto+categoría**, consolidados en una sola celda al construir la tabla resumen. Ejemplo: P-03/Dama = `"0,9 Y 1,1"`; en `BUZO!C3` (Muselina) la fórmula es `=0.9*1.8*F3`, y en `BUZO!J3` (una segunda variante de Muselina, columnas "I") es `=1.1*1.6*F3` — 0.9 y 1.1 son el consumo en dos bloques de tela distintos, no una contradicción.

**Consecuencia para el modelo**: `ConsumoTela` no puede ser solo `Producto × CategoriaConsumo → consumo`. Necesita una tercera dimensión: **tipo/marca de tela específica** (Muselina, Micro, Cottone, Keira Plus, Centauro, Centahelada, Algodón Egipcio, Taslan, Memory, Oxford, Drill Sahara, Mono...). El archivo `franelas cucuta` por sí solo lista **seis tipos de tela distintos** solo para franelas, cada uno con su propio consumo y precio por metro. La tabla `Consumo_Tela_Producto` de 36 filas subestima drásticamente cuántas combinaciones reales existen — el catálogo real de telas (visto también en `Materiales!` del refactor, ~40 telas) es la dimensión que faltaba.

### 17.2 Nuevo: dos líneas de negocio con tarifas propias (NJ Store / NJ Sport)

La hoja `P COSTURA` es una tarifa de mano de obra (corte/planchado/costura por prenda) **duplicada para dos marcas del mismo dueño** — "NJ STORE" y "NJ SPORT" — con precios distintos para la misma operación (ej. costura de franela: $3.000 en NJ Store vs $2.500 en NJ Sport), cada una con sus 3 tramos (Unidad/Docena/Mayor). **No estaba documentado en ninguno de los tres archivos originales ni en `ARCHITECTURE.md`**. Si el negocio opera ambas marcas activamente, el modelo de costos necesita poder escoparse por línea de negocio, no asumir una sola tarifa de mano de obra.

### 17.3 Nuevo: para pedidos grandes personalizados, la "talla" no usa ninguno de los esquemas ya confirmados

`UNI FUTBOL` (uniformes de equipos de fútbol por encargo) y el bloque de fútbol en `PIJAMAS` NO usan Niño/Dama/Caballero/Plus/Unica/Accesorio ni S/M/L/XL/2XL — usan **rangos de edad o tallas literales del cliente**: "TALLAS 6-8-10-12", "CABALLERO S-M-L", "TALLAS 16-S-M", o listas por equipo con talla numérica suelta (Brasil: talla 8/10/12/14, cantidad por talla). Es un **tercer esquema de tallado**, ad hoc por pedido, que ninguno de los dos esquemas ya confirmados (categoría de consumo / talla comercial) cubre. **Implicación para el modelo**: para pedidos por encargo (equipos, uniformes corporativos grandes) puede no ser correcto forzar la talla a un catálogo fijo — el sistema necesita poder aceptar una talla/rango libre asociada a una cantidad, además del catálogo estándar.

### 17.4 Nuevo: el abono del 60% no es universal

`CV TASLAN` (uniformes de fútbol por encargo, 500 unidades) usa **"ABONO 70% ENTREGA 25 DIAS HABILES"**, no 60%. `TAPASOL` sí confirma el 60% con el mismo texto exacto que `Propuesta.xlsm!B24`, para pedidos de catálogo estándar. **Hipótesis, no confirmada**: el 60% es el estándar para catálogo; los pedidos grandes por encargo (con más riesgo/capital inmovilizado) usan un porcentaje mayor — pendiente de confirmar con la dueña si esto es una regla consciente o varió caso a caso.

### 17.5 Nuevo: el "costeo por rendimiento" (precio ÷ unidades que rinde) es mucho más amplio de lo que parecía

No es exclusivo de láminas de sublimación. Aparece igual para: vinil (`LOGO EN VINIL = 14000/210`), marquillas DTF (`MARQUILLA 2 DTF = 41000/256`), telas por corte en lote (`TELA BAHAMAS = 3.7*1.35/80*4900`, un área que rinde 80 piezas), botones por lote (`BOTONES = 0.5*F3`, con nota "*ojo para 30 franelas*" — calibrado para un lote de 30). **Es un patrón de costeo genérico** (precio del insumo comprado a granel ÷ rendimiento en unidades) que aplica a cualquier insumo comprado en lote/lámina/rollo, no una excepción de la sublimación.

### 17.6 Nuevo: si existe precio ligado a color, es al color de IMPRESIÓN, no al color de la prenda

`BOLSAS POP` tiene una tabla de precio explícita por tamaño y **por número de tintas de impresión** (1 tinta vs. 2 tintas, con sobrecosto fijo de +$400 en el precio por 100 unidades). Esto es distinto del hallazgo de la dueña (color del insumo/avío que hay que comprar especial) — pero confirma que "color" sí aparece como variable de costo en al menos un lugar del negocio, con precio ya tabulado (no vía override). Vale la pena diferenciar en el modelo: **color de insumo/avío** (impredecible, vía override — sección 8) vs. **número de colores de una técnica de impresión** (predecible, tarifable de antemano).

### 17.7 Inconsistencias nuevas encontradas en la auditoría completa

| # | Hallazgo | Categoría | Evidencia |
|---|---|---|---|
| 9 | `TULA!B3` (tela para "tulas con delantero sublimado") = `=1*0.9/3`, sin multiplicar por ningún precio ni tasa — a diferencia de cada otra fórmula de tela en las 51 hojas. El resultado (0.3) se suma directamente a costos en USD como si ya fuera un valor monetario. | **A (error de implementación)** — el costo de tela de esa tula específica parece estar subestimado en la práctica a cero | `TULA!B3, B15` |
| 10 | `SUETER` y `HODDI TERRY`: la celda base de costo (`B14`, `B31`) es `=SUM(...)`, **sin dividir por la tasa** — mientras que en 49 de las otras 50 hojas la celda equivalente sí divide por la tasa para dar el costo en USD. Los tramos de margen (`D7=D6/F2`, etc.) sí dividen aparte. | **B (dato/fórmula inconsistente)** | `SUETER!B14`, `HODDI TERRY!B14,B31` |
| 11 | `PAÑOLETA`: el encabezado dice "DELANTAL" pero el nombre de hoja es "PAÑOLETA" y el contenido real incluye un producto no catalogado, "bandana de perro". | **B/F (organización inconsistente, posible producto no documentado en el catálogo)** | `PAÑOLETA!A2, A17` |
| 12 | `BEISBOL` incluye una línea de "GORRA" al final, sin relación con el producto de la hoja. `CV TASLAN`, `MANUBRIO` y la fila `K245:K247` de `fRANELA...` tienen cálculos de una cotización puntual (con abono, ganancia, fecha) pegados directamente en el archivo maestro de precios. | **D (decisión/cálculo manual de una cotización específica, mezclado con el catálogo)** — patrón repetido, no un caso aislado | Varias hojas |

### 17.8 Qué NO cambió tras la auditoría completa

Ninguna de las 19 reglas confirmadas en la sección 6 quedó contradicha por las 51 hojas — al contrario, cada una apareció repetida decenas de veces con distintos productos. La auditoría amplía el modelo (tipo de tela como tercera dimensión, dos líneas de negocio, tallado libre para pedidos por encargo) pero no invalida nada de lo ya confirmado con la dueña.

## 17. Recomendación de siguiente paso

**Actualización 2026-08-17**: 8 de las 10 preguntas críticas originales ya están resueltas (sección 13). Quedan pendientes: el origen exacto de los 7 casos de consumo ambiguo, la reconciliación de los dos esquemas de tramo de volumen, y una decisión de alcance sobre `HOJA DE CALCULO DE PRECIOS AL MAYOR.xlsx`.

No pasar todavía a `/opsx:propose`. Antes:

1. **Decidir el alcance sobre el archivo nuevo** (pregunta crítica #11): ¿auditoría completa de las 51 hojas, o se cierra la Fase 0 con lo ya cubierto y ese archivo queda como referencia puntual? Esto determina si falta una ronda más de este mismo análisis o si ya se puede avanzar.
2. **Si se cierra ya**: verificar puntualmente los 7 casos de consumo ambiguo contra ese archivo (ya hay una hipótesis de trabajo — sección 6, regla #21) antes de construir la tabla de consumo definitiva.
3. **Decidir explícitamente si D-001 (modelo de atributos genéricos) se mantiene, se pospone o se revierte** — este análisis encontró un dominio textil concreto con reglas propias (categoría de consumo por criterio del operador, color como dimensión faltante, desperdicio de material como base real de los tramos de margen) que conviene modelar bien en su forma concreta antes de abstraer.
4. **Agregar Color al alcance del modelo** — no estaba en el pedido original ni en ninguna arquitectura previa, pero es la causa real y frecuente del override, según la propia dueña del negocio.
5. Recién con eso resuelto, `/opsx:propose` para el modelo de datos concreto (catálogo + recetas + cotización de la sección 14), sin tocar aún el override de D-002/D-003 (VBA protegido) hasta que el override por línea esté explícitamente en el alcance.
