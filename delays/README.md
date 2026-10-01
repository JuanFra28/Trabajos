# Estandarizador de delays

`estandarizar_delays.py` convierte el export crudo de delays (el formato de
`delays.xlsx`) al formato estándar de la hoja **Datos_Delays**, que es el que
leen `reporte.py` y el libro `Planes_Delays_2026.xlsx`.

```
delays.xlsx  ──►  estandarizar_delays.py  ──►  Datos_Delays_estandar.xlsx  ──►  reporte.py
 (export crudo)                                 (hoja 1: Datos_Delays)          (HTML)
```

## Uso

1. Copia `estandarizar_delays.py` en la misma carpeta que `reporte.py`, el export
   (`delays.xlsx`) y tu histórico (`Planes_Delays_2026.xlsx`).
2. Ejecuta:

   ```
   python estandarizar_delays.py                          # usa ENTRADA de la configuración
   python estandarizar_delays.py delays.xlsx              # o indica la entrada
   python estandarizar_delays.py delays.xlsx salida.xlsx  # y la salida
   ```

3. Genera el reporte con la salida:

   ```
   python reporte.py Datos_Delays_estandar.xlsx
   ```

   `reporte.py` no necesita cambios: toma la primera hoja (`Datos_Delays`).
   Si pones `EJECUTAR_REPORTE = True`, el estandarizador lo llama al terminar.

Requiere `pandas`, `numpy` y `openpyxl` (los mismos que `reporte.py`).
Opcional: `pip install python-calamine` acelera la lectura de Excel. En las pruebas, el
proceso completo bajó de unos 14 s a unos 8 s.

## De dónde sale cada columna

| Columna estándar | Se obtiene de | Limpieza |
|---|---|---|
| Maquina | `Maquina` | Si viene vacía o no es un código válido (por ejemplo `Even`), se toma de la unidad: `PF5 Converter` → `PF5` |
| Unidad | `PU_Desc` | Espacios de más |
| Fecha | `Fecha_Inicio_Real` | Acepta fecha de Excel, texto `dd/mm/aaaa` o ISO |
| Mes | Fecha | Número de mes |
| Turno | `Turno` | Mayúsculas; `Turno A` → `A` |
| Tipo | `Tipo` (o `Reason_Level1` si falta) | `Averías Eléctricas`, `Electrical`, `Electronico`, `AVERIAS ELECTRICAS`… → **Eléctricas**; las variantes mecánicas → **Mecánicas** |
| Modulo / Cod | `Reason_Level2` (nivel 2 del árbol de razones) | `AS1 - ADHESIVO DE CONSTRUCCION`, `AS1-adhesivo…`, `AS1`, `Adhesivo De Construccion (AS1)`, `PI8 - AS1 - …` → `AS1 — Adhesivo De Construccion` |
| Area | Catálogo de módulos | 113 módulos de tu histórico; los nuevos se asignan con reglas por palabras clave |
| Familia | `Reason_Level3` | Tu histórico manda; los sistemas nuevos se clasifican con reglas |
| Sistema (L3) | `Reason_Level3` | Unifica variantes de mayúsculas, tildes y espacios. Vacío → `(sin dato)` |
| Detalle (L4) | `Reason_Level4` | Quita prefijos de tipo (`Elec - `, `Mech - `, `Minor - `…) y unifica variantes |
| Horas | `Delay` | Si `Delay` viene vacío: `Fecha_Fin_Real − Fecha_Inicio_Real`. Si `Delay` llega en minutos o segundos, se detecta y se convierte |
| Minutos | Horas × 60 | |

Filtros por defecto: solo averías **Eléctricas** y **Mecánicas** desde el
**01/01/2026**, que es lo que contiene tu `Datos_Delays` actual. Se eliminan las
filas duplicadas exactas.

## Importante: el export no trae el módulo

En `delays.xlsx` la columna `Reason_Level12` es **idéntica** a `Reason_Level1`
(las 75.790 filas). El módulo (nivel 2: `AS1`, `OCU`, `FCO`…) no viene en el
archivo y no se puede deducir con fiabilidad del sistema o del detalle. Se
probó y, con ese método, se acierta menos de la mitad de las veces.

Mientras tanto, el script hace lo siguiente:

1. Si el export trae un nivel 2 válido, lo usa. La columna se reconoce como
   `Reason_Level2`, `Nivel2`, etc.
2. Si no lo trae, recupera el módulo desde el **histórico**
   (`HISTORICO = Planes_Delays_2026.xlsx [Datos_Delays]`) cruzando unidad,
   hora de inicio, L3 y L4. Con los datos actuales se recuperan el 100 % de los
   eventos de 2026.
3. Lo que no esté en ninguno de los dos queda como `(sin dato)`. Aparece en la
   hoja *Pendientes* y `reporte.py` lo cuenta como "sin codificar".

**Solución de fondo:** pide o corrige la consulta del export para que traiga
`Reason_Level2`. Con eso los meses nuevos tendrán su módulo sin depender del
histórico.

## Hojas del Excel de salida

| Hoja | Contenido |
|---|---|
| **Datos_Delays** | Las 14 columnas estándar en el mismo orden (A–N), así que las fórmulas de `Delays 80%` y `Fallas por módulo` siguen apuntando bien. Después vienen columnas de trazabilidad: `Fin`, `Año`, `Origen módulo`, `N1/N2/L3/L4 original` y `Observaciones` |
| Resumen | Horas por máquina y tipo, por mes, por área, por familia y los 25 módulos con más horas |
| Calidad | Qué se corrigió y cuántas filas: duplicados, horas calculadas, máquinas rescatadas, filas excluidas por tipo o fecha, origen del módulo y avisos |
| Pendientes | Lo que conviene revisar: sistemas nuevos clasificados por regla, reclasificaciones, módulos fuera de catálogo y eventos sin módulo |
| Unificaciones | Cada texto original de L3 y L4 junto al texto estándar en que quedó y el tipo de cambio aplicado |

## Configuración (inicio del script)

| Parámetro | Por defecto | Para qué |
|---|---|---|
| `ENTRADA`, `SALIDA` | `delays.xlsx`, `Datos_Delays_estandar.xlsx` | Archivos |
| `HISTORICO` | `[("Planes_Delays_2026.xlsx", "Datos_Delays")]` | De dónde recuperar módulos y aprender el catálogo. Puedes agregar salidas anteriores |
| `TIPOS_INCLUIR` | `["Eléctricas", "Mecánicas"]` | `None` = todas las paradas |
| `FECHA_DESDE`, `FECHA_HASTA` | `"2026-01-01"`, `None` | Rango de fechas |
| `QUITAR_PREFIJO_L4` | `True` | `Mech - desgaste de faja` → `Desgaste de faja` |
| `UNIFICAR_VARIANTES` | `True` | `Regulacion electrónica` = `Regulacion Electronica` |
| `RECLASIFICAR_OTROS` | `True` | Si el histórico dejó un sistema en "Otros componentes" pero una regla lo reconoce (`Falla Neumatica` → Neumático), usa la regla |
| `COLUMNAS_EXTRA` | `True` | Columnas de trazabilidad al final |
| `EJECUTAR_REPORTE` | `False` | Llamar a `reporte.py` al terminar |

Correcciones manuales (ganan a todo): `FAMILIAS_MANUAL`, `MODULOS_MANUAL`,
`ALIAS_MODULOS` y `TEXTOS_MANUAL` (sinónimos de L3/L4). En Windows escribe las
rutas como `r"C:\carpeta\archivo.xlsx"` o con `/`.

## Validación con los archivos actuales

Corrido sobre `delays.xlsx` + `Planes_Delays_2026.xlsx`:

- 75.790 filas leídas → 12.870 averías (tu `Datos_Delays` tiene 12.872; la
  diferencia son 2 filas duplicadas exactas).
- Máquina, Unidad, Fecha, Turno, Tipo, Área, Módulo y Cod coinciden al
  **100 %** con tu `Datos_Delays`. Las horas totales también: 3.260,36 h.
- Familia coincide en el 99,5 %. La diferencia son 70 eventos que el histórico
  tenía en "Otros componentes" y que ahora quedan en su familia (por ejemplo
  `Seguridad`, `Red de Comunicación`, `Falla Neumatica`). Están listados en
  *Pendientes*.
- `reporte.py` da el mismo resultado con la salida que con tu `Datos_Delays`:
  9.660 averías, 3.260 h, 349 activos.
- L4: 349 escrituras distintas pasan a 258. Por ejemplo, `Desgaste de faja
  dentada` y `Mech - desgaste de faja dentada` ahora suman juntos.

Si pegas la salida en el libro `Planes_Delays_2026.xlsx`, la lista de fallas de
la hoja *Fallas por módulo* usa los textos L4 viejos (con `Mech - `). Hay que
regenerarla o poner `QUITAR_PREFIJO_L4 = False`.
