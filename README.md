# Powerbi_TFA
# Limpieza y normalización de datos de ventas en Power BI

## Descripción del proyecto

El objetivo del ejercicio fue importar un archivo de ventas generado por un sistema legacy y prepararlo en Power Query antes de construir el modelo de datos.

El archivo original contenía nombres técnicos, valores nulos, filas duplicadas, formatos inconsistentes y datos de clientes mezclados con datos de transacciones. El resultado final incluye dos tablas normalizadas, una relación activa entre ellas y una página de validación en Power BI.

## 1. Carga del archivo

El archivo de Excel se importó mediante el conector **Excel Workbook** de Power BI Desktop.

Se seleccionó la hoja `VENTAS_EXPORT` y se verificó que la vista previa fuera coherente. Posteriormente se eligió **Transform Data** para realizar la limpieza en el Editor de Power Query.

La consulta original se renombró como:

```text
stg_ventas_export
```

Esta consulta se utilizó como tabla intermedia para conservar la estructura original y generar las consultas finales mediante referencias. Su carga al modelo fue deshabilitada porque solamente funciona como origen de las transformaciones.

## 2. Orden de las transformaciones

Las transformaciones se realizaron en el siguiente orden:

1. Importación del archivo de Excel.
2. Selección de la hoja `VENTAS_EXPORT`.
3. Promoción de la primera fila como encabezado.
4. Eliminación de filas completamente vacías.
5. Corrección de los tipos de datos.
6. Corrección de los errores de conversión de `FLG_ACT`.
7. Tratamiento de los valores nulos.
8. Creación de las consultas `dim_clientes` y `fact_ventas`.
9. Selección de las columnas correspondientes a cada tabla.
10. Eliminación de registros duplicados.
11. Estandarización de los valores de `CANAL_VENTA`.
12. Reemplazo de los nombres técnicos por nombres descriptivos.
13. Cálculo de los valores faltantes de `MONTO_TOTAL_VENTA`.
14. Verificación y corrección de los identificadores de cliente y producto.
15. Creación de la relación entre `dim_clientes` y `fact_ventas`.
16. Carga de las tablas al modelo de Power BI.
17. Creación de una página de validación con dos visualizaciones de tipo tabla.

## 3. Renombrado de columnas

Los nombres técnicos del sistema original se reemplazaron por nombres descriptivos en español, utilizando palabras separadas por guiones bajos.

Algunos de los cambios realizados fueron:

| Nombre original | Nombre final           |
| --------------- | ---------------------- |
| `COD_OP`        | `ID_VENTA`             |
| `COD_CLI`       | `ID_CLIENTE`           |
| `NOM_CLI`       | `NOMBRE_CLIENTE`       |
| `MAIL_CLI`      | `EMAIL_CLIENTE`        |
| `TEL_CLI`       | `TELEFONO_CLIENTE`     |
| `CIU_CLI`       | `CIUDAD_CLIENTE`       |
| `PROV_CLI`      | `PROVINCIA_CLIENTE`    |
| `SEG_CLI`       | `SEGMENTO_CLIENTE`     |
| `FLG_ACT`       | `CLIENTE_ACTIVO`       |
| `F_ALTA_CLI`    | `FECHA_ALTA_CLIENTE`   |
| `F_VTA`         | `FECHA_VENTA`          |
| `COD_PROD`      | `ID_PRODUCTO`          |
| `DESC_PROD`     | `NOMBRE_PRODUCTO`      |
| `RUBRO_PROD`    | `CATEGORIA_PRODUCTO`   |
| `CANT`          | `CANTIDAD`             |
| `PU_VTA`        | `PRECIO_UNITARIO`      |
| `DTO_PCT`       | `PORCENTAJE_DESCUENTO` |
| `TOT_VTA`       | `MONTO_TOTAL_VENTA`    |
| `COD_MON`       | `MONEDA`               |
| `CANAL_VTA`     | `CANAL_VENTA`          |

Durante la validación se detectó que `ID_CLIENTE` e `ID_PRODUCTO` habían quedado intercambiados. El problema se corrigió verificando el contenido de las columnas:

* `ID_CLIENTE` debe contener valores como `COD_CLI_005`.
* `ID_PRODUCTO` debe contener códigos como `848`.

Esta revisión permitió crear correctamente la relación entre clientes y ventas.

## 4. Tipos de datos

Los tipos de datos se eligieron según la función de cada columna:

* `ID_VENTA`, `ID_CLIENTE` e `ID_PRODUCTO`: **Text**, porque son identificadores y no deben utilizarse en operaciones matemáticas ni resumirse automáticamente.
* `FECHA_ALTA_CLIENTE` y `FECHA_VENTA`: **Date**, para permitir filtros, agrupaciones y cálculos temporales.
* `CANTIDAD`: **Whole Number**, porque representa unidades completas vendidas.
* `PRECIO_UNITARIO`, `PORCENTAJE_DESCUENTO` y `MONTO_TOTAL_VENTA`: **Decimal Number**, porque participan en cálculos y pueden contener decimales.
* Nombres, ubicaciones, segmento, categoría, moneda y canal de venta: **Text**, porque contienen información descriptiva.
* `CLIENTE_ACTIVO`: **Text**, porque el sistema original utiliza los valores `S` y `N`.

Inicialmente, Power Query intentó convertir `FLG_ACT` a un tipo lógico, lo que generó errores. Se corrigió el paso de tipo de datos y la columna se configuró como texto.

## 5. Tratamiento de valores nulos

Los valores nulos se gestionaron según el significado de cada columna.

### Porcentaje de descuento

Los valores nulos de `PORCENTAJE_DESCUENTO` se reemplazaron por `0`, interpretando que corresponden a operaciones sin descuento.

La transformación se realizó mediante:

```text
Transform → Replace Values
```

La columna se mantuvo como **Decimal Number**.

### Datos de contacto

Los valores nulos de `EMAIL_CLIENTE` y `TELEFONO_CLIENTE` se reemplazaron por etiquetas descriptivas que indican que la información debe completarse, por ejemplo:

```text
Completar email
Completar teléfono
```

Esto permitió conservar a los clientes sin inventar datos de contacto.

### Monto total de venta

Los valores nulos de `MONTO_TOTAL_VENTA` no se reemplazaron por cero, porque eso habría representado incorrectamente operaciones sin facturación.

Se creó una columna personalizada mediante:

```text
Add Column → Custom Column
```

Se utilizó la siguiente fórmula:

```powerquery
if [MONTO_TOTAL_VENTA] = null then
    Number.Round(
        [CANTIDAD] *
        [PRECIO_UNITARIO] *
        (1 - [PORCENTAJE_DESCUENTO]),
        2
    )
else
    [MONTO_TOTAL_VENTA]
```

La fórmula conserva los montos existentes y calcula solamente aquellos que estaban vacíos. La nueva columna se configuró como **Decimal Number**, se eliminó la columna anterior y la columna calculada se renombró como `MONTO_TOTAL_VENTA`.

### Filas vacías

Las filas completamente vacías se eliminaron mediante:

```text
Home → Remove Rows → Remove Blank Rows
```

Estas filas no contenían información útil y su eliminación no implicó perder transacciones válidas.

## 6. Tratamiento de duplicados

El archivo original contenía 945 registros.

### Duplicados en `fact_ventas`

Se utilizó `ID_VENTA` como identificador de cada operación y se aplicó:

```text
Home → Remove Rows → Remove Duplicates
```

Se eliminaron 45 registros duplicados. La tabla quedó formada por 900 ventas únicas.

Los registros eliminados eran repeticiones de operaciones existentes. Mantenerlos habría duplicado los montos facturados y las cantidades vendidas.

### Duplicados en `dim_clientes`

En `dim_clientes` se utilizó `ID_CLIENTE` para conservar un único registro por cliente.

Después de aplicar **Remove Duplicates**, la tabla quedó formada por 119 clientes únicos.

Los duplicados no fueron eliminados de manera indiscriminada: en cada tabla se utilizó el identificador correspondiente para determinar si dos registros representaban la misma entidad.

## 7. Normalización

El archivo original incluía datos descriptivos de clientes repetidos en cada operación de venta. Para evitar esta redundancia, la información se separó en dos tablas.

### Tabla `dim_clientes`

Contiene la información descriptiva del cliente:

* `ID_CLIENTE`
* `NOMBRE_CLIENTE`
* `EMAIL_CLIENTE`
* `TELEFONO_CLIENTE`
* `CIUDAD_CLIENTE`
* `PROVINCIA_CLIENTE`
* `SEGMENTO_CLIENTE`
* `CLIENTE_ACTIVO`
* `FECHA_ALTA_CLIENTE`

Cada cliente aparece una sola vez.

### Tabla `fact_ventas`

Contiene la información de cada operación:

* `ID_VENTA`
* `ID_CLIENTE`
* `FECHA_VENTA`
* `ID_PRODUCTO`
* `NOMBRE_PRODUCTO`
* `CATEGORIA_PRODUCTO`
* `CANTIDAD`
* `PRECIO_UNITARIO`
* `PORCENTAJE_DESCUENTO`
* `MONTO_TOTAL_VENTA`
* `MONEDA`
* `CANAL_VENTA`

El criterio utilizado fue separar los atributos relativamente estables del cliente de los datos que cambian en cada transacción.

## 8. Estandarización del canal de venta

La columna `CANAL_VENTA` contenía valores escritos con diferentes combinaciones de mayúsculas y minúsculas, como:

* `Online`
* `online`
* `ONLINE`
* `Sucursal`
* `SUCURSAL`
* `Telefonico`
* `TELEFONICO`

Se aplicó:

```text
Transform → Format → UPPERCASE
```

Los valores finales quedaron normalizados como:

* `ONLINE`
* `SUCURSAL`
* `TELEFONICO`

Esto evita que Power BI interprete distintas formas de escritura como categorías diferentes.

## 9. Relación entre las tablas

Se creó una relación activa entre:

```text
dim_clientes[ID_CLIENTE]  1 ─── *  fact_ventas[ID_CLIENTE]
```

La configuración utilizada fue:

* **Cardinality:** One to many (1:*)
* **Cross-filter direction:** Single
* **Status:** Active

`dim_clientes` se encuentra del lado uno porque contiene un registro único por cliente. `fact_ventas` se encuentra del lado muchos porque un cliente puede realizar varias operaciones.

## 10. Validación en el reporte

Se creó una única página denominada:

```text
Validación de datos
```

La página contiene dos visualizaciones de tipo **Table**:

1. **Dimensión de clientes:** muestra una selección de campos de `dim_clientes`.
2. **Detalle de ventas:** muestra una selección de campos de `fact_ventas`.

Estas visualizaciones demuestran que las transformaciones se aplicaron correctamente, que las tablas fueron cargadas al modelo y que los datos están disponibles para construir análisis posteriores.

## 11. Evidencias

### Transformaciones en Power Query

La siguiente captura muestra la tabla final en el Editor de Power Query, junto con los pasos aplicados y la validación de las columnas:

![Transformaciones realizadas en Power Query](assets/power-query-transformaciones.png)

### Datos cargados en Power BI

La siguiente captura muestra la página `Validación de datos` con las tablas de clientes y ventas:

![Tablas cargadas en Power BI](assets/validacion-datos-power-bi.png)

## 12. Archivos entregados

El repositorio contiene un único archivo de Power BI actualizado:

```text
Ejercicio_opcional_TFA.pbix
```
No se incluyeron copias duplicadas ni versiones anteriores del archivo `.pbix`.
La estructura final propuesta es:
```text
proyecto-power-bi/
├── README.md
├── Ejercicio_opcional_TFA.pbix
└── assets/
    ├── power-query-transformaciones.png
    └── validacion-datos-power-bi.png
```

## Resultado final

El proceso produjo un modelo formado por:

* 119 clientes únicos en `dim_clientes`.
* 900 operaciones únicas en `fact_ventas`.
* Una relación activa de uno a muchos entre clientes y ventas.
* Una página de validación con dos tablas.
* Columnas con nombres descriptivos y tipos de datos adecuados.
* Valores nulos y duplicados gestionados según criterios técnicos.

El modelo quedó limpio, normalizado y preparado para realizar análisis y visualizaciones en Power BI.
