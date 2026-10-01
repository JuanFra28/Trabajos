# Estandarizador de delays

`estandarizar_delays.py` convierte el export crudo de delays (el formato de
`delays.xlsx`) al **mismo estándar de la hoja `Datos` de
`Analisis_Averias_Mantenimiento.xlsx`**, que es la que lee `reporte.py`.

```
delays.xlsx  ──►  estandarizar_delays.py  ──►  Analisis_Averias_estandar.xlsx  ──►  reporte.py
 (export crudo)          ▲                      (hoja Datos)                         (HTML)
                         │
     Analisis_Averias_Mantenimiento*.xlsx / Planes_Delays_2026*.xlsx
     (referencias: de aquí se copia el módulo de cada evento)
```

## Uso

1. Pon en la misma carpeta: `estandarizar_delays.py`, el export (`delays.xlsx`) y
   al menos una referencia ya clasificada: `Analisis_Averias_Mantenimiento*.xlsx`
   (hoja `Datos`) o `Planes_Delays_2026*.xlsx` (hoja `Datos_Delays`).
2. Ejecuta:

   ```
   python estandarizar_delays.py                          # usa ENTRADA de la configuración
   python estandarizar_delays.py delays.xlsx              # o indica la entrada
   python estandarizar_delays.py delays.xlsx salida.xlsx  # y la salida
   ```

3. Genera el reporte:

   ```
   python reporte.py Analisis_Averias_estandar.xlsx
   ```

   `reporte.py` lee la hoja `Datos` y deja `reporte_averias.html` junto al
   Excel. Si pones `EJECUTAR_REPORTE = True`, el estandarizador lo llama al
   terminar. Más abajo se explica qué muestra el reporte.

Requiere `pandas`, `numpy` y `openpyxl` (los mismos que `reporte.py`).
Opcional: `pip install python-calamine` acelera la lectura de Excel.

## Formato de salida (hoja `Datos`)

Las mismas 13 columnas, en el mismo orden y con el mismo formato que tu archivo:
encabezado azul, fecha `aaaa-mm-dd hh:mm`, horas `0.000`, filtro y primera fila fija.

| Columna | Se obtiene de | Regla |
|---|---|---|
| Maquina | `Maquina` | Si viene vacía o inválida (por ejemplo `Even`), se toma de la unidad: `PF5 Converter` → `PF5` |
| Unidad | `PU_Desc` | Tal cual |
| Fecha | `Fecha_Inicio_Real` | Acepta fecha de Excel, texto `dd/mm/aaaa` o ISO |
| Mes | Fecha | `aaaa-mm` (ej. `2026-01`) |
| Turno | `Turno` | Mayúsculas |
| Tipo | `Tipo` (o `Reason_Level1` si falta) | `Averías Eléctricas`, `Electrical`, `Electronico`… → **Eléctricas**; las variantes mecánicas → **Mecánicas** |
| Area / Modulo / Cod | Nivel 2 del export o tus referencias | `AS1 — Adhesivo De Construccion`, con tu mismo catálogo de 113 módulos |
| Familia | `Reason_Level3` | Tu mismo catálogo; los sistemas nuevos se clasifican con reglas |
| Sistema (L3) | `Reason_Level3` | Tal cual viene |
| Detalle (L4) | `Reason_Level4` | Tal cual viene |
| Horas | `Delay` | Si `Delay` viene vacío: `Fecha_Fin_Real − Fecha_Inicio_Real` |

Por defecto se queda con las averías **Eléctricas** y **Mecánicas** desde el
**01/01/2026**. El orden de las filas es el mismo de tu archivo.

## Filas que se eliminan

- **Sin dato en Sistema (L3) o en Detalle (L4)** (`ELIMINAR_SIN_L3_L4`).
  Para quitar también las que dicen `Libre`, agrega `"Libre"` a
  `SIN_INFO_L3_L4`.
- **Duplicados exactos**: la misma fila repetida en el export
  (`QUITAR_DUPLICADOS`).

Todas quedan listadas en la hoja **Eliminadas**, con su motivo.

## De dónde sale el módulo

El export `delays.xlsx` **no trae el módulo**: la columna `Reason_Level12` es una
copia exacta de `Reason_Level1`. El script resuelve el módulo así:

1. Si el export trae `Reason_Level2` (nivel 2 del árbol de razones), lo usa.
   Reconoce `AS1 - ADHESIVO…`, `AS1`, `Nombre (AS1)` y `PI8 - AS1 - …`.
2. Si no, copia el módulo del mismo evento (unidad + hora de inicio + L3 + L4)
   desde tus referencias (`REFERENCIAS`). Se buscan junto al script, en la
   carpeta actual y junto al export, y manda la versión más reciente.
3. Si el evento no está en ninguna de las dos, queda `(sin dato)`. Esos eventos
   aparecen en la hoja **Pendientes** y en un aviso en pantalla.

Esto significa que los eventos **nuevos** (posteriores a tus referencias) solo
tendrán módulo si el export trae `Reason_Level2`. No es posible deducirlo con
fiabilidad del sistema o del detalle: se probó y acierta menos de la mitad de
las veces.

## Hojas del Excel de salida

| Hoja | Contenido |
|---|---|
| **Datos** | El estándar (13 columnas), listo para `reporte.py` |
| Calidad | Qué se hizo y cuántas filas: horas calculadas, máquinas rescatadas, filas excluidas por tipo o fecha, eliminadas, origen del módulo y avisos |
| Eliminadas | Cada fila eliminada con su motivo (`sin L3 ni L4`, `sin L3`, `sin L4`, `duplicado exacto`) |
| Pendientes | Eventos sin módulo, módulos fuera de catálogo y sistemas nuevos clasificados por regla |
| Resumen | Horas por máquina y tipo, por mes, por área, por familia y los 25 módulos con más horas |

## Configuración (inicio del script)

| Parámetro | Por defecto | Para qué |
|---|---|---|
| `ENTRADA`, `SALIDA` | `delays.xlsx`, `Analisis_Averias_estandar.xlsx` | Archivos |
| `REFERENCIAS` | `Analisis_Averias_Mantenimiento*.xlsx [Datos]`, `Planes_Delays_2026*.xlsx [Datos_Delays]` | De dónde copiar el módulo (se aceptan comodines) |
| `TIPOS_INCLUIR` | `["Eléctricas", "Mecánicas"]` | `None` = todas las paradas |
| `FECHA_DESDE`, `FECHA_HASTA` | `"2026-01-01"`, `None` | Rango de fechas |
| `ELIMINAR_SIN_L3_L4`, `SIN_INFO_L3_L4` | `True`, `["(sin dato)"]` | Qué filas sin L3/L4 se quitan |
| `QUITAR_DUPLICADOS` | `True` | Quita filas repetidas idénticas |
| `QUITAR_PREFIJO_L4` | `False` | `Mech - desgaste de faja` → `Desgaste de faja` |
| `UNIFICAR_VARIANTES` | `False` | `Regulacion electrónica` = `Regulacion Electronica` |
| `RECLASIFICAR_OTROS` | `False` | Saca de "Otros componentes" los sistemas que una regla reconoce |
| `COLUMNAS_EXTRA` | `False` | Agrega columnas de trazabilidad (Minutos, Fin, Origen módulo…) |
| `EJECUTAR_REPORTE` | `False` | Llamar a `reporte.py` al terminar |

Las limpiezas opcionales vienen apagadas para que L3, L4 y Familia salgan igual
que en tu archivo. Si las activas, se agrega la hoja *Unificaciones* con cada
cambio.

Correcciones manuales (ganan a todo): `FAMILIAS_MANUAL`, `MODULOS_MANUAL`,
`ALIAS_MODULOS` y `TEXTOS_MANUAL`. En Windows escribe las rutas como
`r"C:\carpeta\archivo.xlsx"` o con `/`.

## Validación con los archivos actuales

Corrido sobre `delays.xlsx` y comparado con `Analisis_Averias_Mantenimiento-1-2.xlsx`:

- 12.872 filas de averías desde 2026 → se eliminan 140 sin L3/L4 (73 h) y 2
  duplicadas exactas → **12.730 filas**.
- En esas 12.730 filas, comparadas en el mismo orden, las 13 columnas son
  **idénticas** a tu archivo. Eso incluye Módulo, Cod, Área, Familia, L3, L4 y
  Horas. La única excepción son 2 celdas de `Maquina`: tu archivo dice
  `(sin dato)` y el script pone `PF5`/`PF4`, sacados de la unidad.
- Da el mismo resultado usando como referencia el Analisis, el
  `Planes_Delays_2026.xlsx` o ambos.
- `reporte.py` con la salida: 9.543 averías, 3.187 h, 338 activos.

## El reporte (`reporte.py`)

Genera `reporte_averias.html`, una sola página que se abre en cualquier
navegador y responde cuatro preguntas:

| Sección | Qué responde |
|---|---|
| **1. ¿Está funcionando?** | Horas de avería por mes antes y después del mes que elijas (por ejemplo, el mes en que empezó un plan), con un veredicto en palabras y las máquinas o equipos que más bajaron y subieron |
| 2. ¿Dónde se pierden las horas? | Una fila por máquina, con su propio antes y después |
| 3. ¿Qué equipos atender primero? | Máquina · módulo, ordenados por lo que pierden ahora; cómo fallan y qué se repite |
| 4. ¿Qué falla? | Familias de falla y los detalles que más se repiten |

Todo es interactivo: tocar una máquina, un equipo, una familia, un detalle o un
mes filtra la página entera. Arriba se eligen máquinas, tipo, módulo y meses;
abajo se pueden ver y descargar las averías una por una.

El veredicto compara las horas de avería por mes de los dos periodos (solo
meses completos):

- **Mejoró / Empeoró**: la diferencia es mayor que lo que varían normalmente
  los meses dentro de cada periodo (razón de tasas cuasi-Poisson, 95 %).
- **Bajó / Subió**: va en esa dirección, pero todavía puede ser casualidad.
- **Igual**: cambió menos del 10 % (`CAMBIO_MINIMO`).

Dice si cambió, no por qué: no descuenta cambios de producción ni de forma de
registrar.

Además, une como una sola avería los registros del mismo equipo cuando entre
el fin de uno y el inicio del siguiente pasan menos de `UNIR_MINUTOS` (15):
así una parada partida en el cambio de turno no cuenta como varias. También
cuenta juntas las variantes de escritura de L3 y L4 (`UNIFICAR_TEXTOS`).

Configuración principal (inicio del script): `ARCHIVO` (si no existe, busca el
`Analisis_Averias*.xlsx` más reciente), `SALIDA_HTML`, `COMPARAR_DESDE`
(`None` = últimos 3 meses completos; o `"2026-06"`), `MESES_DESPUES`,
`CAMBIO_MINIMO`, `MIN_AVERIAS_VEREDICTO`, `UNIR_MINUTOS`, `FECHA_INICIO` y
`FECHA_FIN`.

Con los datos actuales (01/01 al 03/09/2026), comparando junio–agosto contra
enero–mayo: la planta pasó de 433 a 338 h de avería por mes (−22 %, mejora
confirmada). PI7 mejoró (−34 %); PF4, PI9, PI11 y PI8 bajan sin confirmar;
PF5, PA1 y PF3 suben sin confirmar.
