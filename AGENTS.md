# Contexto esencial para AI agents

Este proyecto genera un reporte Excel de inventario para Brandao Cattle usando archivos fuente de Excel guardados fuera del repo, normalmente en OneDrive. El punto de entrada principal es `inventory_report.py`.

## Objetivo

Crear un solo workbook `.xlsx` con:

- Una hoja principal `Inventory Report` con el inventario actual.
- Hojas adicionales de records diarios por rancho/location para ver entradas, muertes, shipped y balance de inventario dia a dia.

El reporte final se guarda en `Inventory Reports/` dentro del directorio base detectado.

## Archivos principales

- `inventory_report.py`: orquesta todo. Carga los Excel fuente, arma el inventario actual, crea el workbook y agrega las hojas de records.
- `inventory_records.py`: calcula los records diarios y escribe las hojas de records en el workbook.
- `AGENTS.md`: este contexto para futuros agentes.

## Fuentes de datos

Los archivos esperados estan definidos en `RANCH_FILES`:

- `California Inventory`: `California Inventory.xlsx`
- `La Esperanza Ranch`: `Inventory at Dominguez - Guess Cattle.xlsx`
- `Cesar Frias Ranch`: `Inventory at Frias - Guess Cattle.xlsx`

El codigo busca los archivos en `BASE_PATH_CANDIDATES`:

- `C:/Users/cesar/OneDrive/Documentos`
- `/Users/pgcesaare/OneDrive/Documentos`

No editar estos Excel fuente desde el codigo salvo que el usuario lo pida explicitamente.

## Reglas de negocio

- El ownership objetivo es `Brandao Cattle`.
- La hoja principal de inventario actual solo incluye animales con `Status == "Feeding"`.
- El inventario actual agrupa por `Breed`.
- California se separa por `Location`; `Location` es titulo de cada tabla, no columna.
- Los titulos generales por estado son:
  - `California Inventory` -> `California`
  - `La Esperanza Ranch` -> `Washington`
  - `Cesar Frias Ranch` -> `Idaho`
- Washington e Idaho se muestran como titulo general antes de sus tablas de rancho.

## Hoja principal

La hoja `Inventory Report` usa estas columnas:

- `Breed`
- `Quantity`
- `Avg. Price`
- `Avg. DOF`
- `Min Date`
- `Max Date`
- `Total`

La estructura esperada es:

1. Encabezado `BRANDAO CATTLE`, `INVENTORY REPORT`, fecha.
2. Titulo de estado.
3. Tabla(s) del rancho o location.
4. Totales por tabla.
5. Total global.

## Records diarios

Los records diarios usan todo el historial del ownership `Brandao Cattle`, no solo `Feeding`.
Aunque el acumulado se calcula con todo el historial, cada hoja solo muestra los ultimos 30 dias hasta la ultima fecha disponible del record.

Cada hoja de record tiene columnas:

- `Date`
- `Prev. Inventory`
- `Entries`
- `Deaths`
- `Shipped`
- `Inventory`

La logica diaria es:

- `Entries`: conteo por `Date In`.
- `Deaths`: conteo de rows con `Status == "Dead"` por `Death Date`.
- `Shipped`: conteo de rows con `Status == "Shipped"` por `Shipped out date`.
- `Inventory`: acumulado de `entries - deaths - shipped`.
- `Prev. Inventory`: inventario del dia anterior.

Si una fila tiene `Status == "Shipped"` pero no tiene `Shipped out date`, el codigo usa `Date In` como fecha de salida para que no quede inflando el inventario activo.

Las hojas de records esperadas actualmente son:

- `Gold Star Cattle Record`
- `Vazquez Calf Ranch Record`
- `La Esperanza Ranch Record`
- `Cesar Frias Ranch Record`

California usa las mismas locations activas detectadas en el inventario actual para crear sus hojas de records.

## Estilo Excel

Mantener estilo similar en todas las hojas:

- Sin gridlines.
- Margenes: left/right `0.25`, top/bottom `0.75`, header/footer `0.3`.
- `fitToWidth = 1`.
- Las hojas de records usan `fitToHeight = 1` para imprimir en una sola pagina cuando sea posible.
- Altura de filas `18`.
- El contenido del area impresa debe quedar alineado verticalmente al centro sin cambiar la alineacion horizontal existente.
- Encabezados con fuente bold.
- Bordes inferiores grises en headers de tabla.
- Las filas de datos llevan separadores horizontales gris claro.
- Los rows `TOTAL` llevan borde superior gris para separar la tabla del total.
- Formatos de fecha `mm/dd/yyyy`.
- Cantidades enteras con `#,##0`.
- Valores monetarios con `$#,##0.00`.

## Ejecucion

Comando principal:

```bash
python3 inventory_report.py
```

Esto debe crear un archivo como:

```text
BRANDAO CATTLE INVENTORY REPORT mm.dd.yyyy.xlsx
```

en:

```text
<BASE_PATH>/Inventory Reports/
```

## Verificacion recomendada

Despues de cambios, correr:

```bash
python3 -m py_compile inventory_report.py inventory_records.py
```

Para probar sin escribir en OneDrive, llamar `generate_inventory_report` con un output temporal:

```bash
python3 -c "from pathlib import Path; import inventory_report as ir; ir.generate_inventory_report(ir.inventory_assignments, Path('/private/tmp/brandao_inventory_test.xlsx')); print('ok')"
```

Validar que el workbook tenga:

- `Inventory Report`
- una hoja de record por rancho/location esperado
- mismos margenes en todas las hojas
- cierre de `Inventory` en records igual al inventario `Feeding` actual de esa hoja

## Cuidado al modificar

- No romper la compatibilidad con `gold_star_inv`, `vazquez_calf_ranch_inv`, `la_esperanza_inv`, `cesar_frias_ranch_inv` y `frias_ranch_inv`; pueden usarse desde otros scripts.
- Evitar duplicar lectura de Excel si se puede reutilizar `ranch_dataframes`.
- Si se agregan nuevos ranchos, actualizar `RANCH_FILES` y, si aplica, `RANCH_SECTION_TITLES`.
- Si se agregan nuevos estados o locations, cuidar que los nombres de sheets no pasen 31 caracteres y no tengan caracteres invalidos de Excel.
