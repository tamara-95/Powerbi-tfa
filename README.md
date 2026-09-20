# Powerbi_TFA
# Preparación y transformación de datos de ventas en Power BI

## Objetivo

El objetivo del ejercicio fue importar un archivo de ventas proveniente de un sistema legacy y transformarlo en Power Query para obtener información limpia, consistente y organizada antes de construir el modelo de datos en Power BI.

## 1. Carga del archivo

El archivo se importó desde Power BI Desktop utilizando el conector **Excel Workbook**.

Antes de comenzar la transformación se verificó que la vista previa fuera coherente y que la primera fila contuviera los encabezados correspondientes. Luego se seleccionó **Transform Data** para abrir el Editor de Power Query.

La consulta original se renombró como `stg_ventas_export` y se utilizó como consulta de preparación. A partir de ella se crearon las tablas finales mediante consultas de referencia.

## 2. Transformaciones realizadas y orden aplicado

Las transformaciones se realizaron en el siguiente orden:

1. Importación del archivo de Excel.
2. Selección de la hoja `VENTAS_EXPORT`.
3. Promoción de la primera fila como encabezado.
4. Eliminación de filas completamente vacías.
5. Revisión y corrección de los tipos de datos.
6. Corrección de la columna `FLG_ACT`, que inicialmente generaba errores al ser interpretada como un valor lógico.
7. Gestión de los valores nulos.
8. Creación de las consultas `dim_clientes` y `fact_ventas`.
9. Selección de las columnas correspondientes a cada tabla.
10. Eliminación de registros duplicados.
11. Normalización de los valores de `canal_venta`.
12. Reemplazo de los nombres técnicos por nombres descriptivos.
13. Cálculo de los valores nulos de `monto_total_venta`.
14. Revisión final de errores, valores vacíos y tipos de datos.

## 3. Renombrado de columnas

Los nombres técnicos del sistema original se reemplazaron por nombres descriptivos en español, utilizando palabras separadas por guiones bajos.

Algunos ejemplos son:

* `COD_CLI` → `id_cliente`
* `NOM_CLI` → `nombre_cliente`
* `MAIL_CLI` → `email_cliente`
* `F_ALTA_CLI` → `fecha_alta_cliente`
* `COD_OP` → `id_venta`
* `F_VTA` → `fecha_venta`
* `COD_PROD` → `id_producto`
* `DESC_PROD` → `nombre_producto`
* `RUBRO_PROD` → `categoria_producto`
* `CANT` → `cantidad`
* `PU_VTA` → `precio_unitario`
* `DTO_PCT` → `porcentaje_descuento`
* `TOT_VTA` → `monto_total_venta`
* `COD_MON` → `moneda`
* `CANAL_VTA` → `canal_venta`

Este cambio facilita la lectura de los datos y el mantenimiento posterior del modelo.

## 4. Tipos de datos seleccionados

Los tipos de datos se definieron según la función de cada columna.

* Los identificadores, como `id_cliente`, `id_venta` e `id_producto`, se configuraron como **Text** porque representan códigos y no valores destinados a operaciones matemáticas. Esto también evita que Power BI intente sumarlos.
* `fecha_alta_cliente` y `fecha_venta` se configuraron como **Date** para permitir filtros temporales, agrupaciones por período y cálculos entre fechas.
* `cantidad` se configuró como **Whole Number** porque representa unidades completas vendidas.
* `precio_unitario`, `porcentaje_descuento` y `monto_total_venta` se configuraron como **Decimal Number** porque pueden contener decimales y participan en cálculos.
* Los campos descriptivos, como nombres, ciudad, provincia, categoría, moneda y canal de venta, se configuraron como **Text**.
* `cliente_activo` se configuró como **Text** porque el sistema original utiliza los valores `S` y `N`. Inicialmente Power Query intentó convertir esta columna a un tipo lógico, lo que generó errores. El problema se resolvió modificando su tipo a texto.

## 5. Resolución de valores nulos

Los valores nulos no se trataron todos de la misma manera. Primero se analizó qué representaba la ausencia de información en cada columna.

### Porcentaje de descuento

Los valores nulos de `porcentaje_descuento` se reemplazaron por `0`, considerando que representan operaciones en las que no se aplicó descuento.

La transformación se realizó seleccionando la columna y utilizando:

**Transform → Replace Values**

Se reemplazó `null` por `0` y la columna se configuró como **Decimal Number**.

### Correo electrónico y teléfono

Los valores nulos de `email_cliente` y `telefono_cliente` se reemplazaron por `No informado`.

Se utilizó:

**Transform → Replace Values**

Este criterio permitió conservar a los clientes sin inventar direcciones de correo electrónico o números telefónicos.

### Monto total de venta

Los valores nulos de `monto_total_venta` no se reemplazaron por cero, ya que esto habría indicado incorrectamente que las operaciones no generaron ingresos.

Para recuperar los importes faltantes se creó una columna personalizada desde:

**Add Column → Custom Column**

Se utilizó la siguiente lógica:

```powerquery
if [monto_total_venta] = null then
    Number.Round(
        [cantidad] *
        [precio_unitario] *
        (1 - [porcentaje_descuento]),
        2
    )
else
    [monto_total_venta]
```

La fórmula conserva los montos existentes y calcula solamente los registros que tenían un valor nulo.

La nueva columna se configuró como **Decimal Number**. Luego se eliminó la columna anterior y la columna calculada se renombró como `monto_total_venta`.

### Filas completamente vacías

Las filas completamente vacías se eliminaron mediante:

**Home → Remove Rows → Remove Blank Rows**

Estas filas no contenían información útil y su eliminación no implicaba perder operaciones válidas.

## 6. Resolución de duplicados

El archivo original contenía 945 filas. Para evitar que una misma operación se contabilizara más de una vez, se revisaron los identificadores de las tablas finales.

### Duplicados en `fact_ventas`

En `fact_ventas` se seleccionó `id_venta` y se utilizó:

**Home → Remove Rows → Remove Duplicates**

El campo `id_venta` se utilizó porque identifica de manera única cada operación.

Se eliminaron 45 registros duplicados y la tabla quedó formada por 900 ventas únicas. Los registros eliminados eran repeticiones de operaciones existentes, por lo que conservarlos habría duplicado los montos y las cantidades vendidas.

### Duplicados en `dim_clientes`

En `dim_clientes` se seleccionó `id_cliente` y se aplicó:

**Home → Remove Rows → Remove Duplicates**

El campo `id_cliente` se utilizó porque cada cliente debe aparecer una sola vez en la dimensión.

Luego de la eliminación de duplicados, la tabla quedó formada por 119 clientes únicos.

Los duplicados no se eliminaron de manera indiscriminada: se utilizaron los identificadores correspondientes para determinar cuándo dos registros representaban al mismo cliente o a la misma operación.

## 7. Normalización de la estructura

El archivo original contenía datos del cliente mezclados con datos de las transacciones. Para reducir la repetición de información, los datos se separaron en dos consultas.

### Tabla `dim_clientes`

Contiene los atributos descriptivos y relativamente estables del cliente:

* `id_cliente`
* `nombre_cliente`
* `email_cliente`
* `telefono_cliente`
* `ciudad_cliente`
* `provincia_cliente`
* `segmento_cliente`
* `cliente_activo`
* `fecha_alta_cliente`

Cada cliente aparece una sola vez.

### Tabla `fact_ventas`

Contiene la información correspondiente a cada operación:

* `id_venta`
* `id_cliente`
* `fecha_venta`
* `id_producto`
* `nombre_producto`
* `categoria_producto`
* `cantidad`
* `precio_unitario`
* `porcentaje_descuento`
* `monto_total_venta`
* `moneda`
* `canal_venta`

El criterio utilizado fue separar los datos que describen al cliente de los datos que cambian en cada transacción.

El campo `id_cliente` se mantuvo en ambas tablas para permitir una relación de uno a muchos:

```text
dim_clientes[id_cliente]  1 ─── *  fact_ventas[id_cliente]
```

Esta separación reduce la repetición de información personal, mejora la organización del modelo y facilita el análisis posterior.

## 8. Estandarización del canal de venta

La columna `canal_venta` contenía valores escritos de diferentes formas, por ejemplo:

* `Online`
* `online`
* `ONLINE`
* `Sucursal`
* `SUCURSAL`
* `Telefonico`
* `TELEFONICO`

Para evitar que Power BI interpretara estos valores como cat
