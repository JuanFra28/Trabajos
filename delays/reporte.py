"""
===============================================================================
 AVERIAS (DELAYS): ¿DONDE ESTAN LOS PROBLEMAS Y SI LO QUE HACEMOS FUNCIONA?  v6
===============================================================================
 Reporte HTML interactivo de una sola página. Responde cuatro preguntas:

   1. ¿Está funcionando?  Horas de avería por mes antes y después del mes que
      elijas, con un veredicto en palabras: mejoró, empeoró o igual.
   2. ¿Dónde se pierden las horas?  Una fila por máquina.
   3. ¿Qué equipos atender primero?  Máquina · módulo, cómo falla y qué es lo
      que más se repite.
   4. ¿Qué falla?  Familias de falla y detalles más repetidos.

 Todo se filtra con un clic: tocar una máquina, un equipo, una familia o un mes
 filtra la página entera. Arriba se eligen máquinas, tipo, módulo y meses.

 USO:
   python reporte.py                      (usa ARCHIVO; si no existe, busca el
                                           Excel de averías más reciente)
   python reporte.py otro_archivo.xlsx    (o pasa la ruta como argumento)

 Requiere: pandas numpy openpyxl
===============================================================================
"""

import glob
import json
import math
import os
import re
import sys
import unicodedata
import warnings
from datetime import datetime

import numpy as np
import pandas as pd

warnings.filterwarnings("ignore")

# =============================================================================
# CONFIGURACION   <<<<<<  AJUSTA AQUI
# =============================================================================
ARCHIVO = "Analisis_Averias_estandar.xlsx"   # en Windows: r"C:\carpeta\archivo.xlsx"
HOJA = "Datos"                 # si no existe, la primera que empiece con "Datos"
SALIDA_HTML = None             # None: reporte_averias.html junto al Excel

PLANTA = "Planta Santa Clara"
TITULO = "Averías: dónde están los problemas y si lo que hacemos funciona"

# --- ¿Está funcionando?  (antes contra después)
COMPARAR_DESDE = None        # "2026-06": primer mes del "después" (por ejemplo,
                             # el mes en que empezó una acción). None: los
                             # últimos MESES_DESPUES meses completos.
MESES_DESPUES = 3
CAMBIO_MINIMO = 10           # % por debajo del cual se dice "igual"
MIN_AVERIAS_VEREDICTO = 10   # con menos averías se dice "pocos datos"

# --- Datos
UNIR_MINUTOS = 15            # une registros del mismo equipo cuando entre el fin
                             # de uno y el inicio del siguiente hay menos de N min
FECHA_INICIO = None          # "2026-02-01" para recortar el inicio
FECHA_FIN = None             # "2026-08-31" para recortar el final
EXCLUIR_SIN_CODIFICAR = False
UNIFICAR_TEXTOS = True       # "Regulacion Electronica" = "Regulación electrónica"
MAX_FILAS_TABLA = 30000      # averías embebidas en el reporte

# --- Formato de números y de CSV
IDIOMA = "es-PE"             # formato de miles y decimales del reporte
CSV_ESTILO = "punto"         # "punto": separa con , y decimal .   (Perú, EE.UU.)
                             # "coma" : separa con ; y decimal ,   (España, Argentina)

# --- Nombres completos de módulos.
#     Se leen solos de la columna "Modulo" del Excel ("OCU — Outer Cover Unwind").
#     Este diccionario solo sirve para corregir alguno puntual.
MODULOS = {
    # "OCU": "Outer Cover Unwind",
}
ARCHIVO_MODULOS = "modulos.csv"      # opcional: columnas código,nombre


# =============================================================================
# 1. ¿FUNCIONA?  (mismas fórmulas que usa el navegador)
# =============================================================================
DIAS_MES = 30.44
# Cuantil 97.5 % de la t de Student, por grados de libertad
T975 = [None, 12.706, 4.303, 3.182, 2.776, 2.571, 2.447, 2.365, 2.306, 2.262, 2.228,
        2.201, 2.179, 2.160, 2.145, 2.131, 2.120, 2.110, 2.101, 2.093, 2.086,
        2.080, 2.074, 2.069, 2.064, 2.060, 2.056, 2.052, 2.048, 2.045, 2.042]


def t975(gl):
    return T975[gl] if gl < len(T975) else 1.96 + 2.4 / gl


def comparar(y, e, corte):
    """Antes contra después.

    y: horas de avería de cada mes completo; e: días observados de ese mes;
    corte: posición del primer mes del "después".
    Compara las horas por mes de los dos periodos y mide si la diferencia es
    mayor que lo que varían normalmente los meses dentro de cada periodo
    (razón de tasas cuasi-Poisson, 95 %). Así un mes malo aislado no basta
    para decir que algo empeoró o mejoró.

    estado: MEJORO / EMPEORO (diferencia confirmada), BAJO / SUBIO (va en esa
    dirección pero todavía puede ser variación normal), IGUAL (cambio menor a
    CAMBIO_MINIMO %). None si no hay al menos 2 meses en cada periodo."""
    y = np.asarray(y, float)
    e = np.asarray(e, float)
    k = len(y)
    if corte < 2 or k - corte < 2:
        return None
    y1, y2, e1, e2 = y[:corte].sum(), y[corte:].sum(), e[:corte].sum(), e[corte:].sum()
    if y1 + y2 <= 0 or e1 <= 0 or e2 <= 0:
        return None
    r1, r2 = y1 / e1, y2 / e2
    mu = np.r_[r1 * e[:corte], r2 * e[corte:]]
    ok = mu > 0
    phi = max(float(((y[ok] - mu[ok]) ** 2 / mu[ok]).sum()) / (k - 2), 1e-9)
    n, p0 = y1 + y2, e2 / (e1 + e2)
    z = (y2 - n * p0) / math.sqrt(phi * n * p0 * (1 - p0))
    cambio = (r2 / r1 - 1) * 100 if r1 > 0 else math.inf
    lim = t975(k - 2)
    if z < -lim:
        estado = "MEJORO"
    elif z > lim:
        estado = "EMPEORO"
    elif abs(cambio) < CAMBIO_MINIMO:
        estado = "IGUAL"
    else:
        estado = "BAJO" if cambio < 0 else "SUBIO"
    return dict(antes=r1 * DIAS_MES, despues=r2 * DIAS_MES, cambio=cambio,
                z=float(z), estado=estado)


# =============================================================================
# 2. CARGA  (tolerante a cambios de columnas y a data nueva)
# =============================================================================
def _norm(s):
    s = unicodedata.normalize("NFKD", str(s))
    s = "".join(c for c in s if not unicodedata.combining(c))
    return re.sub(r"[^a-z0-9]", "", s.lower())


ALIAS = {
    "maquina": ["maquina", "maq", "machine", "linea"],
    "unidad": ["unidad", "unit"],
    "fecha": ["fechahora", "fecha", "date", "timestamp", "inicio"],
    "turno": ["turno", "shift"],
    "tipo": ["tipofalla", "tipo", "type"],
    "area": ["area", "zona"],
    "modulo": ["modulo", "module"],
    "cod": ["cod", "codigo", "code"],
    "familia": ["familia", "family"],
    "l3": ["sistemal3", "sistema", "l3"],
    "l4": ["detallel4", "detalle", "l4"],
    "horas": ["horas", "hrs", "duracion", "hours", "downtime"],
}
SIN = "(sin dato)"
PAT_SIN_COD = r"LIBRE|SIN CLASIF|NO CLASIF|SIN DATO|OTROS / SIN MAPEAR"
PREFIJO = re.compile(r"^\s*(elec|elect|mech|mec|minor|menor)\s*[-–—:]\s*", re.I)
# Detalles (L4) que dicen que hubo que intervenir, pero no qué falló. Se comparan
# en minúsculas, sin tildes y sin prefijos como "Elec -".
GENERICO = re.compile(
    r"^(regulacion( electronica| mecanica)?|regulacion ?/ ?calibracion|falla general|"
    r"otros?( ajustes| cambios)?( ?/ ?ajustes)?|ajuste electronico|"
    r"ajuste (de )?otros seteos|reparacion mecanica|reemplazo de componentes?|"
    r"cambio( de)? componentes?|libre|sin dato|\(sin dato\)|sin clasificar|"
    r"no clasificado|varios)( \(.*\))?$")


def mapear_columnas(df):
    """Encuentra las columnas aunque cambien de nombre o de orden."""
    norm = {_norm(c): c for c in df.columns}
    usadas, out = set(), {}
    for destino, opciones in ALIAS.items():
        hallado = None
        for o in opciones:                       # coincidencia exacta
            if o in norm and norm[o] not in usadas:
                hallado = norm[o]
                break
        if hallado is None:                      # coincidencia parcial
            for k, orig in norm.items():
                if orig in usadas:
                    continue
                if any(k.startswith(o) or o in k for o in opciones):
                    hallado = orig
                    break
        if hallado is not None:
            out[destino] = hallado
            usadas.add(hallado)
    return out


def _txt(serie):
    s = serie.astype(str).str.strip()
    return s.replace({"nan": SIN, "": SIN, "None": SIN, "NaT": SIN, "-": SIN})


def _clave_txt(s):
    s = unicodedata.normalize("NFKD", str(s))
    s = "".join(c for c in s if not unicodedata.combining(c))
    return re.sub(r"\s+", " ", s).strip().lower()


def es_generico(v):
    return bool(GENERICO.match(_clave_txt(PREFIJO.sub("", str(v)))))


def unificar(serie):
    """Variantes de escritura del mismo texto -> la forma más usada."""
    cuenta = serie.value_counts()
    limpio = {v: (PREFIJO.sub("", v).strip() or v) for v in cuenta.index}
    votos = {}
    for v, n in cuenta.items():
        k = _clave_txt(limpio[v])
        votos.setdefault(k, {})
        votos[k][limpio[v]] = votos[k].get(limpio[v], 0) + int(n)
    canon = {k: max(d.items(), key=lambda kv: kv[1])[0] for k, d in votos.items()}
    mapa = {v: canon[_clave_txt(limpio[v])] for v in cuenta.index}
    return serie.map(mapa), sum(1 for v, c in mapa.items() if v != c)


def buscar_archivo(ruta):
    """Si la ruta configurada no existe, toma el Excel de averías más reciente
    de la carpeta actual o de la del script."""
    if ruta and os.path.exists(ruta):
        return ruta
    cand = set()
    for carpeta in (os.getcwd(), os.path.dirname(os.path.abspath(__file__))):
        for patron in ("Analisis_Averias*.xlsx", "*verias*.xlsx"):
            cand.update(p for p in glob.glob(os.path.join(carpeta, patron))
                        if not os.path.basename(p).startswith("~$"))
    return max(cand, key=os.path.getmtime) if cand else ruta


def cargar(ruta):
    if ruta.lower().endswith((".csv", ".txt")):
        df = pd.read_csv(ruta, sep=None, engine="python")
    else:
        xl = pd.ExcelFile(ruta)
        hojas = xl.sheet_names
        # la hoja pedida; si no está, la primera que empiece igual ("Datos_Delays")
        hoja = HOJA if HOJA in hojas else next(
            (h for h in hojas if _norm(h).startswith(_norm(HOJA))), hojas[0])
        df = xl.parse(hoja)
    df.columns = [str(c).strip() for c in df.columns]
    m = mapear_columnas(df)

    faltan = [k for k in ("maquina", "fecha", "horas") if k not in m]
    if faltan:
        raise SystemExit(f"\n  No encuentro columnas obligatorias: {faltan}\n"
                         f"  Columnas del archivo: {list(df.columns)}\n")

    d = pd.DataFrame(index=df.index)
    d["Maquina"] = _txt(df[m["maquina"]]).str.upper()
    d["Fecha"] = pd.to_datetime(df[m["fecha"]], errors="coerce")
    d["Horas"] = pd.to_numeric(df[m["horas"]], errors="coerce").fillna(0).clip(lower=0)

    for campo, col in [("Unidad", "unidad"), ("Modulo", "modulo"), ("Cod", "cod"),
                       ("Familia", "familia"), ("Tipo", "tipo"), ("Area", "area"),
                       ("L3", "l3"), ("L4", "l4"), ("Turno", "turno")]:
        d[campo] = _txt(df[m[col]]) if col in m else SIN

    # Máquina vacía pero la unidad sí la trae: "PF5 Converter" -> PF5
    hueco = d["Maquina"].str.contains(re.escape(SIN), case=False, regex=True)
    if hueco.any():
        rescate = d.loc[hueco, "Unidad"].str.extract(r"^([A-Za-z]{1,4}\d{1,3})",
                                                     expand=False)
        d.loc[hueco & rescate.notna(), "Maquina"] = rescate.dropna().str.upper()

    # "CS — Central System" -> código CS + nombre Central System
    part = d["Modulo"].astype(str).str.split(r"\s+[—–-]\s+", n=1, regex=True)
    cod_txt = part.str[0].str.strip().str.upper()
    nom_txt = part.str[1].fillna("").str.strip() if part.str.len().max() > 1 else ""
    d["ModCod"] = d["Cod"].str.upper() if "cod" in m else cod_txt
    d["ModCod"] = d["ModCod"].replace({SIN.upper(): SIN})
    d["ModNom"] = nom_txt
    d["Modulo"] = d["ModCod"]
    d["activo"] = d["Maquina"] + " · " + d["Modulo"]

    unificados = 0
    if UNIFICAR_TEXTOS:
        for c in ("L3", "L4"):
            d[c], n = unificar(d[c])
            unificados += n

    d = d.dropna(subset=["Fecha"]).sort_values("Fecha").reset_index(drop=True)
    d.attrs["unificados"] = unificados
    return d


def cargar_nombres_modulos(df=None):
    """Prioridad: nombres del Excel  <  dict MODULOS  <  modulos.csv"""
    nombres = {}
    if df is not None and "ModNom" in df.columns:
        vistos = (df[df["ModNom"].astype(str).str.len() > 0]
                  .groupby("ModCod")["ModNom"].agg(lambda x: x.mode().iloc[0]))
        nombres.update({str(k).upper(): str(v) for k, v in vistos.items()})
    nombres.update(MODULOS)
    if os.path.exists(ARCHIVO_MODULOS):
        try:
            mm = pd.read_csv(ARCHIVO_MODULOS, dtype=str).fillna("")
            cols = list(mm.columns)
            if len(cols) >= 2:
                for _, r in mm.iterrows():
                    cod, nom = str(r[cols[0]]).strip().upper(), str(r[cols[1]]).strip()
                    if cod and nom:
                        nombres[cod] = nom
        except Exception as e:
            print(f"        Aviso: no pude leer {ARCHIVO_MODULOS} ({e})")
    return {k.strip().upper(): v for k, v in nombres.items()}


# =============================================================================
# 3. CALIDAD DEL DATO, PERIODO Y CONSOLIDACION
# =============================================================================
def diagnostico(df):
    """Medidas de calidad sobre los registros tal como vienen."""
    tot = max(1e-9, float(df["Horas"].sum()))
    d = dict(n_reg=len(df), horas_total=tot)
    mask = pd.Series(False, index=df.index)
    for c in ["L3", "L4", "Familia", "Modulo"]:
        mask |= df[c].str.upper().str.contains(PAT_SIN_COD, na=False, regex=True)
    d["mask_sin_cod"] = mask
    d["horas_sin_cod"] = float(df.loc[mask, "Horas"].sum())
    d["pct_sin_cod_h"] = d["horas_sin_cod"] / tot * 100
    peor = (df.assign(_hs=df["Horas"].where(mask, 0)).groupby("Maquina")
              .agg(hs=("_hs", "sum"), h=("Horas", "sum")).reset_index())
    peor["pct"] = peor["hs"] / peor["h"].where(peor["h"] > 0) * 100
    peor = peor[peor["h"] > tot * .02].sort_values("pct", ascending=False)
    d["peor_cod"] = ((peor.iloc[0]["Maquina"], float(peor.iloc[0]["pct"]))
                     if len(peor) and peor.iloc[0]["pct"] > 0 else None)
    gmap = {v: es_generico(v) for v in df["L4"].unique()}
    gen = df["L4"].map(gmap).astype(bool)
    d["pct_generico_h"] = float(df.loc[gen, "Horas"].sum() / tot * 100)
    corta = df["Horas"] * 60 < 2
    d["pct_cortas"] = float(corta.mean() * 100)
    d["pct_cortas_h"] = float(df.loc[corta, "Horas"].sum() / tot * 100)
    return d


def periodo(df):
    """Inicio y fin observados, y días cubiertos de cada mes."""
    f0 = df["Fecha"].min()
    f1 = (df["Fecha"] + pd.to_timedelta(df["Horas"], unit="h")).max()
    ini = f0.to_period("M").to_timestamp() if f0.day <= 5 else f0.normalize()
    fin = f1
    # el recorte manual nunca extiende el periodo más allá de los datos: un mes
    # sin datos contaría como un mes sin averías y falsearía la comparación
    if FECHA_INICIO:
        ini = max(ini, pd.Timestamp(FECHA_INICIO))
    if FECHA_FIN:
        fin = min(fin, pd.Timestamp(FECHA_FIN) + pd.Timedelta(days=1))
    cov, m = [], ini.to_period("M")
    while m.to_timestamp() < fin:
        a = max(ini, m.to_timestamp())
        b = min(fin, (m + 1).to_timestamp())
        cov.append([m.year * 12 + m.month - 1,
                    round((b - a).total_seconds() / 86400, 3), m.days_in_month])
        m += 1
    return ini, fin, cov


def consolidar(df, minutos=UNIR_MINUTOS):
    """Une los registros que son la misma parada: los del mismo equipo cuando
    entre el fin de uno y el inicio del siguiente hay menos de N minutos.
    Así una parada partida en el cambio de turno cuenta como una sola.
    Sistema, detalle, turno, etc. se toman del registro más largo."""
    d = df.sort_values(["activo", "Fecha"]).reset_index(drop=True)
    fin = d["Fecha"] + pd.to_timedelta(d["Horas"], unit="h")
    fin_prev = fin.groupby(d["activo"]).cummax().groupby(d["activo"]).shift()
    hueco = (d["Fecha"] - fin_prev).dt.total_seconds() / 60
    nuevo = hueco.isna() | (hueco >= minutos)
    d["_g"] = nuevo.cumsum()
    largo = d.loc[d.groupby("_g")["Horas"].idxmax()].set_index("_g")
    agg = d.groupby("_g").agg(Fecha=("Fecha", "min"), Horas=("Horas", "sum"),
                              n_sub=("Horas", "size"))
    cols = ["Maquina", "Unidad", "Modulo", "ModCod", "ModNom", "Familia", "Tipo",
            "Area", "L3", "L4", "Turno", "Cod", "activo"]
    out = agg.join(largo[[c for c in cols if c in d.columns]])
    out = out.sort_values("Fecha").reset_index(drop=True)
    info = dict(unidos=len(d) - len(out),
                continuas=int(((hueco.abs() < 1) & ~nuevo).sum()))
    return out, info


def meses_completos(cov):
    """Meses con al menos la mitad de los días observados (para comparar)."""
    return [(k, dd) for k, dd, dm in cov if dd >= 0.5 * dm]


def corte_inicial(cov):
    """Mes en que empieza el "después" por defecto (clave año*12+mes-1)."""
    ks = [k for k, _ in meses_completos(cov)]
    if not ks:
        return None
    if COMPARAR_DESDE:
        p = pd.Period(COMPARAR_DESDE, "M")
        k = p.year * 12 + p.month - 1
        if k in ks:
            return k
        print(f"        Aviso: {COMPARAR_DESDE} no es un mes completo de los datos; "
              f"se usan los últimos {MESES_DESPUES} meses.")
    i = len(ks) - MESES_DESPUES if len(ks) >= MESES_DESPUES + 2 else len(ks) // 2
    return ks[max(i, 0)]


def resumen_comparacion(av, cov, corte):
    """Lo mismo que muestra el navegador sin filtros, para la consola."""
    vm = meses_completos(cov)
    ks, e = [k for k, _ in vm], np.array([dd for _, dd in vm], float)
    if corte not in ks:
        return None, {}
    c = ks.index(corte)
    mk = av["Fecha"].dt.year * 12 + av["Fecha"].dt.month - 1
    dentro = mk.isin(ks)

    def serie(mask):
        s = av[mask].groupby(mk[mask])["Horas"].sum()
        return s.reindex(ks, fill_value=0).to_numpy(float)

    planta = comparar(serie(pd.Series(True, index=av.index)), e, c)
    maqs = {}
    for m in sorted(av["Maquina"].unique()):
        msk = av["Maquina"] == m
        if (msk & dentro).sum() >= MIN_AVERIAS_VEREDICTO:
            maqs[m] = comparar(serie(msk), e, c)
    return planta, maqs


# =============================================================================
# 4. PAYLOAD PARA EL NAVEGADOR
# =============================================================================
DIMS = ["Maquina", "Unidad", "Modulo", "Familia", "Tipo", "Area", "Turno", "L3", "L4"]


def payload(av, nombres, dq, limite=MAX_FILAS_TABLA):
    """Codifica las averías como índices enteros contra diccionarios.
    th = horas desde el inicio del periodo."""
    d = av.sort_values("Fecha").copy()
    truncado = len(d) > limite
    if truncado:
        d = d.nlargest(limite, "Horas").sort_values("Fecha")
    d["th"] = (d["Fecha"] - dq["ini"]).dt.total_seconds() / 3600.0

    dic, idx = {}, {}
    for c in DIMS:
        vals = sorted(d[c].astype(str).unique())
        dic[c], idx[c] = vals, {v: i for i, v in enumerate(vals)}
    cols = [d[c].astype(str).map(idx[c]).to_numpy() for c in DIMS]
    hrs, th, ns = (d["Horas"].round(4).to_numpy(), d["th"].round(3).to_numpy(),
                   d["n_sub"].to_numpy())
    filas = [[int(c[i]) for c in cols] + [float(hrs[i]), float(th[i]), int(ns[i])]
             for i in range(len(d))]
    gen = [1 if es_generico(v) else 0 for v in dic["L4"]]

    return dict(
        dic={k.lower(): v for k, v in dic.items()},
        modnom=[nombres.get(m.upper(), "") for m in dic["Modulo"]],
        gen=gen, rows=filas, cov=dq["cov"],
        t0=dq["ini"].strftime("%Y-%m-%d %H:%M:%S"),
        fin=dq["fin"].strftime("%Y-%m-%d %H:%M:%S"),
        meta=dict(
            planta=PLANTA, titulo=TITULO,
            emitido=datetime.now().strftime("%d/%m/%Y"),
            nReg=int(dq["n_reg"]), nFinal=int(len(av)), nEmbed=len(d),
            truncado=bool(truncado), horasTotal=float(av["Horas"].sum()),
            corte=dq["corte"], cambioMin=CAMBIO_MINIMO, minVer=MIN_AVERIAS_VEREDICTO,
            pctSinCodH=round(dq["pct_sin_cod_h"], 1),
            peorCod=(list(dq["peor_cod"]) if dq.get("peor_cod") else None),
            pctGenericoH=round(dq["pct_generico_h"], 1),
            idioma=IDIOMA,
            csvSep=(";" if CSV_ESTILO == "coma" else ","),
            csvDec=("," if CSV_ESTILO == "coma" else "."),
        ))


# =============================================================================
# 5. ESTILOS
# =============================================================================
CSS = r"""
:root{
  --tinta:#0B2430; --tinta2:#1B3B4A; --gris:#5C7480; --gris2:#8199A3;
  --linea:#D5DEE0; --linea2:#E9EEEF; --fondo:#EDF1F1; --panel:#FFFFFF; --panel2:#F6F9F9;
  --rojo:#B5352B; --rojoBg:#FBEDEB; --rojoBd:#EED2CE;
  --naranja:#B0560F; --naranjaBg:#FDF3EA; --naranjaBd:#F0D7C2;
  --ambar:#9C6B00; --ambarF:#E0A200; --ambarBg:#FCF5E4; --ambarBd:#EDDCB6;
  --verde:#17714B; --verdeBg:#E9F4EF; --verdeBd:#C6E1D4;
  --verde2:#3E8E68; --verde2Bg:#F1F8F4; --verde2Bd:#D5E9DE;
  --azul:#10607F; --azul3:#93C6D8; --azulBg:#EAF2F6; --antes:#B9C8CF;
  --r:5px; --sombra:0 10px 34px rgba(11,36,48,.14);
  --sans:"IBM Plex Sans","Segoe UI",system-ui,-apple-system,Roboto,Arial,sans-serif;
}
*,*::before,*::after{box-sizing:border-box}
html{-webkit-text-size-adjust:100%}
body{margin:0;background:var(--fondo);color:var(--tinta2);font-family:var(--sans);
 font-size:15.5px;line-height:1.55;font-feature-settings:"tnum" 1,"lnum" 1}
h1,h2,h3{margin:0;font-weight:600;color:var(--tinta);letter-spacing:-.011em}
p{margin:0 0 10px}
.wrap{max-width:1120px;margin:0 auto;padding:0 24px}
.oculto{display:none !important}
.suave{color:var(--gris)}
.num{font-variant-numeric:tabular-nums}

/* ---------- cabecera ---------- */
.rail{background:var(--tinta);color:#fff}
.rail .wrap{display:flex;align-items:flex-end;justify-content:space-between;gap:20px;
 padding-top:22px;padding-bottom:18px;flex-wrap:wrap}
.rail .planta{color:var(--azul3);font-size:14px;font-weight:500;margin:0 0 5px}
.rail h1{color:#fff;font-size:24px;line-height:1.2;max-width:34ch}
.railmeta{color:#A9C4CE;font-size:13.5px;text-align:right;line-height:1.7}
.railmeta b{color:#fff;font-weight:600}
.railmeta button{background:none;border:1px solid rgba(255,255,255,.3);color:#D7E6EC;
 padding:5px 12px;border-radius:var(--r);font-size:13px;cursor:pointer;font-family:inherit;
 margin-top:6px}
.railmeta button:hover{background:rgba(255,255,255,.1)}

/* ---------- filtros ---------- */
.filtros{background:var(--panel);border-bottom:1px solid var(--linea);position:sticky;top:0;
 z-index:40;box-shadow:0 2px 10px rgba(11,36,48,.06)}
.filtros .wrap{padding-top:11px;padding-bottom:11px}
.frow{display:flex;gap:8px;align-items:center;flex-wrap:wrap}
.frow+.frow{margin-top:8px}
.fetq{font-size:12.5px;color:var(--gris);font-weight:600;min-width:62px}
.sep{width:1px;height:24px;background:var(--linea);margin:0 6px}
.chip{appearance:none;font-family:inherit;background:var(--panel);border:1px solid var(--linea);
 color:var(--tinta2);padding:5px 11px;border-radius:var(--r);font-size:13.8px;font-weight:500;
 cursor:pointer;line-height:1.3}
.chip:hover{border-color:var(--azul);background:var(--azulBg)}
.chip[aria-pressed="true"]{background:var(--azul);border-color:var(--azul);color:#fff}
.chip:focus-visible{outline:2px solid var(--tinta);outline-offset:1px}
.pill{appearance:none;font-family:inherit;background:var(--azulBg);border:1px solid #C6DEE8;
 color:var(--azul);padding:3px 10px;border-radius:20px;font-size:13px;cursor:pointer}
.pill:hover{background:#DCEBF2}
select{font-family:inherit;font-size:13.8px;padding:6px 9px;border:1px solid var(--linea);
 border-radius:var(--r);background:var(--panel);color:var(--tinta2);max-width:230px}
input[type=search]{font-family:inherit;font-size:13.8px;padding:6px 10px;
 border:1px solid var(--linea);border-radius:var(--r);background:var(--panel);color:var(--tinta2)}
select:focus,input:focus{outline:2px solid var(--azul);outline-offset:-1px}
.btn{appearance:none;font-family:inherit;background:var(--panel);border:1px solid var(--azul);
 color:var(--azul);padding:5px 12px;border-radius:var(--r);font-size:13.6px;font-weight:500;
 cursor:pointer}
.btn:hover{background:var(--azulBg)}
.btn:disabled{opacity:.4;cursor:default}
.estado{font-size:13.2px;color:var(--gris);margin-left:auto}
.estado b{color:var(--tinta)}

/* ---------- secciones ---------- */
main{padding:22px 0 50px}
.sec{background:var(--panel);border:1px solid var(--linea);border-radius:var(--r);
 padding:20px 22px 22px;margin-bottom:18px}
.sechead{display:flex;align-items:center;gap:12px;flex-wrap:wrap;margin-bottom:6px}
.sechead .n{display:inline-grid;place-items:center;width:28px;height:28px;border-radius:50%;
 background:var(--tinta);color:#fff;font-weight:600;font-size:14px;flex:none}
.sechead h2{font-size:21px}
.sechead .der{margin-left:auto;display:flex;gap:8px;align-items:center;font-size:14px;
 color:var(--gris);flex-wrap:wrap}
.lead{color:var(--gris);font-size:14.4px;margin:0 0 14px;max-width:80ch}
.sec h3{font-size:16px;margin:4px 0 8px}
.cols{display:grid;gap:22px}
.cols.c2{grid-template-columns:minmax(0,1fr) minmax(0,1fr)}

/* ---------- veredicto ---------- */
.ver{border-radius:var(--r);padding:18px 20px;border:1px solid var(--linea);
 border-left:6px solid var(--gris2);background:var(--panel2);margin-bottom:16px}
.ver .t{font-size:24px;font-weight:600;color:var(--tinta);line-height:1.2}
.ver .s{font-size:15.5px;color:var(--tinta2);margin-top:6px;max-width:86ch}
.ver .cifras{display:flex;gap:30px;flex-wrap:wrap;margin-top:14px}
.ver .cifra .v{font-size:26px;font-weight:600;color:var(--tinta);line-height:1.1}
.ver .cifra .v i{font-style:normal;font-size:14px;color:var(--gris);font-weight:500;margin-left:3px}
.ver .cifra .k{font-size:13px;color:var(--gris);margin-top:3px}
.ver.MEJORO{border-left-color:var(--verde);background:var(--verdeBg);border-color:var(--verdeBd)}
.ver.MEJORO .t{color:var(--verde)}
.ver.BAJO{border-left-color:var(--verde2);background:var(--verde2Bg);border-color:var(--verde2Bd)}
.ver.SUBIO{border-left-color:var(--naranja);background:var(--naranjaBg);border-color:var(--naranjaBd)}
.ver.EMPEORO{border-left-color:var(--rojo);background:var(--rojoBg);border-color:var(--rojoBd)}
.ver.EMPEORO .t{color:var(--rojo)}
.cambios h3{font-size:15px;margin-bottom:6px}
.lista{margin:0;padding:0;list-style:none}
.lista li{display:flex;gap:10px;justify-content:space-between;align-items:center;
 padding:7px 0;border-bottom:1px solid var(--linea2);font-size:14.2px;cursor:pointer}
.lista li:last-child{border-bottom:0}
.lista li:hover .nom{color:var(--azul);text-decoration:underline}
.lista .nom{font-weight:600;color:var(--tinta)}
.lista .val{color:var(--gris);font-size:13.4px;white-space:nowrap;text-align:right}

/* ---------- etiquetas de estado ---------- */
.badge{display:inline-block;padding:2px 9px;border-radius:20px;font-size:12.4px;
 font-weight:600;border:1px solid;white-space:nowrap;line-height:1.5}
.e-MEJORO{color:#fff;background:var(--verde);border-color:var(--verde)}
.e-BAJO{color:var(--verde);background:var(--verde2Bg);border-color:var(--verde2Bd)}
.e-IGUAL{color:var(--gris);background:var(--panel2);border-color:var(--linea)}
.e-SUBIO{color:var(--naranja);background:var(--naranjaBg);border-color:var(--naranjaBd)}
.e-EMPEORO{color:#fff;background:var(--rojo);border-color:var(--rojo)}
.e-NA{color:var(--gris2);background:none;border-color:transparent;font-weight:500}
.gen{color:var(--ambar);font-size:12.4px;white-space:nowrap}

/* ---------- gráficos ---------- */
.chart{position:relative;width:100%}
.chart svg{display:block;width:100%;height:auto;overflow:visible}
.chart .hit{cursor:pointer}
.tip{position:fixed;z-index:200;display:none;pointer-events:none;width:250px;
 background:var(--tinta);color:#DCE9ED;border-radius:var(--r);padding:10px 12px;
 font-size:12.9px;line-height:1.45;box-shadow:var(--sombra)}
.tip .t{font-weight:600;color:#fff;font-size:13.4px;margin-bottom:6px;padding-bottom:5px;
 border-bottom:1px solid rgba(255,255,255,.17)}
.tip .r{display:flex;justify-content:space-between;gap:12px;padding:2px 0}
.tip .r span{color:#9FBAC4}
.tip .r b{color:#fff;font-weight:600}
.tip .q{color:#9FBAC4;font-size:11.9px;margin-top:6px;padding-top:6px;
 border-top:1px solid rgba(255,255,255,.15)}
.leg{display:flex;flex-wrap:wrap;gap:14px;align-items:center;margin-top:8px;font-size:13px;
 color:var(--gris)}
.leg i{width:12px;height:12px;border-radius:2px;display:inline-block;margin-right:6px;
 vertical-align:-1px}

/* ---------- tablas ---------- */
.tablaBox{border:1px solid var(--linea);border-radius:var(--r);overflow:hidden}
.tablaTop{display:flex;gap:10px;align-items:center;padding:8px 12px;
 border-bottom:1px solid var(--linea);background:var(--panel2);flex-wrap:wrap}
.tablaTop .tt{font-size:13.4px;color:var(--gris)}
.tablaTop .der{margin-left:auto;display:flex;gap:8px;align-items:center}
.scroll{overflow-x:auto;max-width:100%}
table{border-collapse:collapse;width:100%;font-size:14px}
thead th{background:var(--panel2);text-align:left;padding:9px 12px;font-weight:600;
 color:var(--tinta);border-bottom:1px solid var(--linea);white-space:nowrap;font-size:13px}
th.s{cursor:pointer;user-select:none}
th.s:hover{background:#ECF2F3}
th.s span.ar{opacity:0;margin-left:5px;font-size:9px}
th.s.asc span.ar,th.s.desc span.ar{opacity:.9}
th.s.asc span.ar::after{content:"\25B2"}
th.s.desc span.ar::after{content:"\25BC"}
tbody td{padding:8px 12px;border-bottom:1px solid var(--linea2);white-space:nowrap}
tbody tr:last-child td{border-bottom:0}
tbody tr.click{cursor:pointer}
tbody tr.click:hover td{background:var(--azulBg)}
td.r,th.r{text-align:right}
td.b{font-weight:600;color:var(--tinta)}
td.w{white-space:normal;min-width:170px;max-width:270px}
td.w2{white-space:normal;min-width:110px;max-width:150px}
b.eq{display:block;color:var(--tinta)}
.sub{display:block;font-size:12.6px;color:var(--gris);max-width:200px;overflow:hidden;
 text-overflow:ellipsis}
.leyv{font-size:13px;color:var(--gris);margin-top:8px;max-width:100ch}
.leyv b{color:var(--tinta2);font-weight:600}
.barcell{position:relative;display:block;min-width:110px;text-align:right}
.barcell i{position:absolute;left:0;top:50%;transform:translateY(-50%);height:16px;
 background:var(--azul3);border-radius:2px;opacity:.5}
.barcell span{position:relative}
.pager{display:flex;gap:9px;align-items:center;padding:8px 12px;font-size:13.2px;
 color:var(--gris);border-top:1px solid var(--linea);flex-wrap:wrap}
.vacio{padding:26px 20px;text-align:center;color:var(--gris);font-size:14.4px}
.vacio b{display:block;color:var(--tinta);font-size:15.2px;margin-bottom:4px}

/* ---------- notas ---------- */
details.bloque{background:var(--panel);border:1px solid var(--linea);border-radius:var(--r);
 padding:15px 22px;margin-bottom:18px}
details.bloque>summary{cursor:pointer;font-size:16.5px;font-weight:600;color:var(--tinta);
 list-style:none}
details.bloque>summary::-webkit-details-marker{display:none}
details.bloque>summary::before{content:"\25B8";display:inline-block;margin-right:9px;
 color:var(--azul);transition:transform .15s}
details.bloque[open]>summary::before{transform:rotate(90deg)}
details.bloque>summary span{font-size:14px;font-weight:400;color:var(--gris);margin-left:6px}
details.bloque[open]>summary{margin-bottom:12px}
details.bloque ul{margin:0;padding-left:20px;font-size:14.4px;max-width:90ch}
details.bloque li{margin-bottom:7px}
.nota{border:1px solid var(--ambarBd);border-left:4px solid var(--ambarF);
 background:var(--ambarBg);padding:12px 16px;border-radius:0 var(--r) var(--r) 0;
 font-size:14.2px;margin-top:14px}
.nota b{color:var(--tinta)}

footer{color:var(--gris);padding:6px 0 30px;font-size:13px}

@media (max-width:900px){
  .cols.c2{grid-template-columns:minmax(0,1fr)}
  .filtros{position:static}
}
@media (max-width:640px){
  body{font-size:15px}
  .wrap{padding:0 14px}
  .rail h1{font-size:20px}
  .railmeta{text-align:left}
  .sec{padding:16px 14px}
  .ver .t{font-size:20px}
  .ver .cifra .v{font-size:22px}
  .fetq{min-width:100%}
  .estado{margin-left:0}
  .sechead .der{margin-left:0}
}
@media (prefers-reduced-motion:reduce){*{transition:none !important}}
@media print{
  body{background:#fff}
  .filtros,.railmeta button,.pager,.tablaTop .der{display:none !important}
  .sec,details.bloque{break-inside:avoid;border-color:#CCC}
  .rail{background:#fff !important}
  .rail h1,.railmeta b{color:#000 !important}
  .rail .planta,.railmeta{color:#444 !important}
}
"""


# =============================================================================
# 6. MOTOR DEL NAVEGADOR
# =============================================================================
JS = r"""(function(){
'use strict';
var D = __DATA__;
var M = D.meta;

/* ============================ constantes ============================ */
var F = {MAQ:0, UNI:1, MOD:2, FAM:3, TIP:4, ARE:5, TUR:6, L3:7, L4:8, H:9, T:10, N:11, MK:12};
var DIMS = [
  {k:'maquina', f:0, t:'Máquina'},
  {k:'tipo',    f:4, t:'Tipo'},
  {k:'modulo',  f:2, t:'Módulo'},
  {k:'familia', f:3, t:'Familia'},
  {k:'l4',      f:8, t:'Detalle'}
];
var MESES = ['ene','feb','mar','abr','may','jun','jul','ago','set','oct','nov','dic'];
var MESL  = ['enero','febrero','marzo','abril','mayo','junio','julio','agosto',
             'setiembre','octubre','noviembre','diciembre'];
var VNOM = {MEJORO:'Mejoró', BAJO:'Bajó', IGUAL:'Igual', SUBIO:'Subió', EMPEORO:'Empeoró'};
var VAYU = {MEJORO:'Bajó más de lo que varía normalmente de un mes a otro: mejora confirmada',
            BAJO:'Va a la baja, pero todavía puede ser variación normal (sin confirmar)',
            IGUAL:'Cambió menos del '+M.cambioMin+' %',
            SUBIO:'Va al alza, pero todavía puede ser variación normal (sin confirmar)',
            EMPEORO:'Subió más de lo que varía normalmente de un mes a otro: empeoramiento confirmado'};
var VORD = {EMPEORO:0, SUBIO:1, IGUAL:2, BAJO:3, MEJORO:4};
var CNOM = {CRITICO:'Seguido y largo', CRONICO:'Seguido, pero corto',
            AGUDO:'Pocas veces, pero largo', LEVE:'Poco y corto'};
var CTXT = {
  CRITICO:'Falla más veces que el promedio y cada parada dura más que el promedio. '+
          'Hay que evitar que se repita y también reparar más rápido.',
  CRONICO:'Falla más veces que el promedio, pero se arregla rápido. '+
          'El foco es que no vuelva a fallar (causa raíz).',
  AGUDO:'Falla menos que el promedio, pero cada parada es larga. '+
        'El foco es reparar más rápido (repuesto a mano, acceso, procedimiento).',
  LEVE:'Falla poco y se arregla rápido.'};
var T975 = [null,12.706,4.303,3.182,2.776,2.571,2.447,2.365,2.306,2.262,2.228,
            2.201,2.179,2.160,2.145,2.131,2.120,2.110,2.101,2.093,2.086,
            2.080,2.074,2.069,2.064,2.060,2.056,2.052,2.048,2.045,2.042];
var DIAS_MES = 30.44;

/* ============================ estado ============================ */
var S = {sel:{maquina:[], tipo:[], modulo:[], familia:[], l4:[]},
         q:'', m0:null, m1:null, corte:null, tab:{}};
var MESKEYS = [], COV = {};
// VM: meses completos del filtro (para comparar); CI: posición del primer mes
// del "después" dentro de VM; DIAS: días cubiertos; RANGO: fechas del filtro
var VM = [], CI = -1, DIAS = 1, RANGO = {ini:0, fin:0};

/* ============================ utilidades ============================ */
function el(id){ return document.getElementById(id); }
function esc(s){ return String(s==null?'':s).replace(/[&<>"]/g, function(c){
  return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]; }); }
function fnum(v,d){ if(v==null||isNaN(v)) return '—';
  return Number(v).toLocaleString(M.idioma||'es-PE',
    {minimumFractionDigits:d, maximumFractionDigits:d}); }
function fh(v){ if(v==null||isNaN(v)) return '—';
  return fnum(v, v>=100?0:(v>=10?1:2)); }
function fh1(v){ if(v==null||isNaN(v)) return '—'; return fnum(v, v>=100?0:1); }
function fmin(h){ if(h==null||isNaN(h)) return '—'; var m=h*60; return fnum(m, m<10?1:0); }
function fpc(v,d){ return fnum(v, d==null?1:d)+' %'; }
function plural(n,a,b){ return n===1?a:b; }
function fechaObj(th){ return new Date(D.t0ms + th*3600000); }
function fechaHora(th){ var d=fechaObj(th), p=function(x){return (x<10?'0':'')+x;};
  return d.getUTCFullYear()+'-'+p(d.getUTCMonth()+1)+'-'+p(d.getUTCDate())+' '+
         p(d.getUTCHours())+':'+p(d.getUTCMinutes()); }
function fechaTxt(ms){ var d=new Date(ms);
  return d.getUTCDate()+' de '+MESL[d.getUTCMonth()]+' de '+d.getUTCFullYear(); }
function mesKeyMs(ms){ var d=new Date(ms); return d.getUTCFullYear()*12 + d.getUTCMonth(); }
function mesTxt(k){ return MESES[k%12]+' '+String(Math.floor(k/12)).slice(2); }
function mesCorto(k){ return MESES[k%12]+' '+Math.floor(k/12); }
function mesLargo(k){ return MESL[k%12]+' de '+Math.floor(k/12); }
function mesMs(k){ return Date.UTC(Math.floor(k/12), k%12, 1); }
function rangoMeses(a,b){ return a===b ? mesTxt(a) : mesTxt(a)+' a '+mesTxt(b); }
function modLabel(i){ var n=D.modnom[i]; return n ? D.dic.modulo[i]+' — '+n : D.dic.modulo[i]; }
function dimTexto(k,i){ return k==='modulo' ? modLabel(i) : D.dic[k][i]; }
function actLabel(m,mo){ return D.dic.maquina[m]+' · '+D.dic.modulo[mo]; }
function actFull(m,mo){ return D.dic.maquina[m]+' · '+modLabel(mo); }
function suma(a){ var s=0; for(var i=0;i<a.length;i++) s+=a[i]; return s; }
function sumaH(rows){ var s=0; for(var i=0;i<rows.length;i++) s+=rows[i][F.H]; return s; }
function clamp(v,a,b){ return v<a?a:(v>b?b:v); }
function corta(s,n){ s=String(s); return s.length>n ? s.slice(0,n-1)+'…' : s; }
function mayus(s){ return s.charAt(0).toUpperCase()+s.slice(1); }

/* ============================ ¿funciona? ============================ */
function t975(gl){ return gl < T975.length ? T975[gl] : 1.96 + 2.4/gl; }
/* Antes contra después (igual que en el script Python).
   y: horas de avería de cada mes completo, e: días observados de ese mes,
   c: posición del primer mes del "después". Mide si la diferencia es mayor
   que lo que varían normalmente los meses dentro de cada periodo
   (razón de tasas cuasi-Poisson, 95 %): un mes malo aislado no alcanza. */
function comparar(y, e, c){
  var k = y.length, i;
  if(c < 2 || k-c < 2) return null;
  var y1=0, y2=0, e1=0, e2=0;
  for(i=0;i<k;i++){ if(i<c){ y1+=y[i]; e1+=e[i]; } else { y2+=y[i]; e2+=e[i]; } }
  if(!(y1+y2>0) || !(e1>0) || !(e2>0)) return null;
  var r1 = y1/e1, r2 = y2/e2, chi = 0, mu;
  for(i=0;i<k;i++){ mu = (i<c ? r1 : r2)*e[i]; if(mu>0) chi += (y[i]-mu)*(y[i]-mu)/mu; }
  var phi = Math.max(chi/(k-2), 1e-9), n = y1+y2, p0 = e2/(e1+e2);
  var z = (y2 - n*p0)/Math.sqrt(phi*n*p0*(1-p0)), lim = t975(k-2);
  var cambio = r1>0 ? (r2/r1-1)*100 : Infinity, estado;
  if(z < -lim) estado = 'MEJORO';
  else if(z > lim) estado = 'EMPEORO';
  else if(Math.abs(cambio) < M.cambioMin) estado = 'IGUAL';
  else estado = cambio < 0 ? 'BAJO' : 'SUBIO';
  return {antes:r1*DIAS_MES, despues:r2*DIAS_MES, cambio:cambio, z:z, estado:estado};
}
function flecha(c){
  if(!isFinite(c.cambio)) return '▲ nuevo';
  return (c.cambio<0?'▼ ':'▲ ')+fnum(Math.abs(c.cambio),0)+' %';
}
function badgeV(c){
  if(!c) return '<span class="badge e-NA" title="Hacen falta al menos '+M.minVer+
    ' averías y 2 meses completos antes y después">pocos datos</span>';
  return '<span class="badge e-'+c.estado+'" title="'+esc(VAYU[c.estado])+'">'+VNOM[c.estado]+
    (c.estado==='IGUAL' ? '' : ' '+flecha(c))+'</span>';
}
var LEYENDA_V = '<b>Mejoró</b> / <b>Empeoró</b>: cambio confirmado, mayor que lo que varía '+
  'normalmente de un mes a otro. <b>Bajó</b> / <b>Subió</b>: va en esa dirección, pero todavía '+
  'puede ser casualidad. <b>Igual</b>: cambió menos del '+M.cambioMin+' %.';
function antesDespues(c){
  return c ? fh1(c.antes)+' → '+fh1(c.despues) : '—';
}

/* ============================ filtrado ============================ */
function pasa(r, excepto){
  var d, sel;
  for(var j=0;j<DIMS.length;j++){
    d = DIMS[j];
    if(d.k===excepto) continue;
    sel = S.sel[d.k];
    if(sel.length && sel.indexOf(r[d.f])===-1) return false;
  }
  if(r[F.MK] < S.m0 || r[F.MK] > S.m1) return false;
  if(S.q){
    var t = (D.dic.l3[r[F.L3]]+' '+D.dic.l4[r[F.L4]]+' '+modLabel(r[F.MOD])+' '+
             D.dic.maquina[r[F.MAQ]]+' '+D.dic.familia[r[F.FAM]]).toLowerCase();
    if(t.indexOf(S.q)===-1) return false;
  }
  return true;
}
function filtra(excepto){
  var out=[], R=D.rows, i;
  for(i=0;i<R.length;i++) if(pasa(R[i], excepto)) out.push(R[i]);
  return out;
}
function corteDefecto(){
  var k = VM.length;
  if(k < 4) return Math.floor(k/2);
  if(M.corte!=null && S.m0===MESKEYS[0] && S.m1===MESKEYS[MESKEYS.length-1]){
    for(var i=0;i<k;i++) if(VM[i].k===M.corte) return i;
  }
  return k >= 5 ? k-3 : 2;
}
function calculaVentana(){
  VM = []; DIAS = 0; RANGO = {ini:null, fin:null};
  MESKEYS.forEach(function(k){
    if(k < S.m0 || k > S.m1) return;
    var c = COV[k];
    DIAS += c.d;
    if(RANGO.ini===null) RANGO.ini = c.a;
    RANGO.fin = c.b;
    if(c.d >= 0.5*c.dm) VM.push({k:k, e:c.d});
  });
  DIAS = Math.max(DIAS, 1/24);
  CI = -1;
  if(S.corte!=null) VM.forEach(function(m,i){ if(m.k===S.corte) CI = i; });
  if(CI < 2 || VM.length-CI < 2){ CI = corteDefecto(); S.corte = VM[CI] ? VM[CI].k : null; }
}
function hayFiltros(){
  var n = 0;
  DIMS.forEach(function(d){ n += S.sel[d.k].length; });
  return n + (S.q?1:0) + ((S.m0!==MESKEYS[0] || S.m1!==MESKEYS[MESKEYS.length-1])?1:0);
}

/* ============================ agrupacion ============================ */
function agrupa(rows, keyFn){
  var map = Object.create(null), ord = [], i, r, k, g, enVM = {};
  VM.forEach(function(m){ enVM[m.k] = 1; });
  for(i=0;i<rows.length;i++){
    r = rows[i]; k = keyFn(r); g = map[k];
    if(!g){ g = map[k] = {key:k, n:0, h:0, r0:r, mods:Object.create(null),
                          l4:Object.create(null), mh:Object.create(null), nVM:0};
            ord.push(g); }
    g.n++; g.h += r[F.H];
    g.mods[r[F.MOD]] = (g.mods[r[F.MOD]]||0) + r[F.H];
    g.l4[r[F.L4]] = (g.l4[r[F.L4]]||0) + 1;
    g.mh[r[F.MK]] = (g.mh[r[F.MK]]||0) + r[F.H];
    if(enVM[r[F.MK]]) g.nVM++;
  }
  var e = VM.map(function(m){ return m.e; });
  for(i=0;i<ord.length;i++){
    g = ord[i];
    g.mttr = g.h/g.n;
    g.cmp = g.nVM >= M.minVer ?
      comparar(VM.map(function(m){ return g.mh[m.k]||0; }), e, CI) : null;
  }
  return ord;
}
function porActivo(rows){
  return agrupa(rows, function(r){ return r[F.MAQ]+'|'+r[F.MOD]; }).map(function(g){
    g.maq=g.r0[F.MAQ]; g.mod=g.r0[F.MOD];
    g.label=actLabel(g.maq,g.mod); g.full=actFull(g.maq,g.mod); return g; });
}
function porMaquina(rows){
  return agrupa(rows, function(r){ return String(r[F.MAQ]); }).map(function(g){
    g.maq=g.r0[F.MAQ]; g.label=D.dic.maquina[g.maq]; g.full=g.label; return g; });
}
function porDim(rows, dimKey){
  var f = DIMS.filter(function(d){return d.k===dimKey;})[0].f;
  return agrupa(rows, function(r){ return String(r[f]); }).map(function(g){
    g.idx=g.r0[f]; g.label=dimTexto(dimKey,g.idx); g.full=g.label; return g; });
}
function porHoras(a){ return a.slice().sort(function(x,y){ return y.h-x.h; }); }
function masRepetido(g){
  var mx=-1, s0=null;
  Object.keys(g.l4).forEach(function(x){ if(g.l4[x]>mx){ mx=g.l4[x]; s0=+x; } });
  return {l4:s0, n:mx};
}
/* Cómo falla (Jack-Knife de Knights 2001): frecuencia contra el promedio de
   fallas por equipo, duración contra la duración promedio de una falla. */
function comoFalla(acts){
  if(!acts.length) return;
  var sn=0, sh=0;
  acts.forEach(function(g){ sn+=g.n; sh+=g.h; });
  var ln = sn/acts.length, lm = sh/sn;
  acts.forEach(function(g){
    g.q = (g.n>ln && g.mttr>lm) ? 'CRITICO' : (g.n>ln) ? 'CRONICO'
        : (g.mttr>lm) ? 'AGUDO' : 'LEVE'; });
}
function filtraEquipo(k){
  var p = String(k).split(/[|:]/);
  S.sel.maquina = [+p[0]]; S.sel.modulo = [+p[1]]; render(); arriba();
}
function arriba(){
  var s = el('s1'); if(s && window.scrollY > s.offsetTop) window.scrollTo(0, Math.max(0, s.offsetTop-120));
}

/* ============================ tooltip ============================ */
var TIP;
function tipHtml(titulo, filas, pie){
  return '<div class="t">'+esc(titulo)+'</div>'+filas.map(function(f){
    return '<div class="r"><span>'+f[0]+'</span><b>'+f[1]+'</b></div>'; }).join('')+
    (pie ? '<div class="q">'+pie+'</div>' : '');
}
function tipShow(html, ev){ if(!TIP) return; TIP.innerHTML = html; TIP.style.display='block'; tipMove(ev); }
function tipMove(ev){
  if(!TIP || TIP.style.display==='none') return;
  var w = TIP.offsetWidth || 250, h = TIP.offsetHeight || 140;
  var x = ev.clientX + 16, y = ev.clientY + 16;
  if(x + w > window.innerWidth - 8) x = ev.clientX - w - 14;
  if(y + h > window.innerHeight - 8) y = Math.max(6, ev.clientY - h - 12);
  TIP.style.left = x+'px'; TIP.style.top = y+'px';
}
function tipHide(){ if(TIP) TIP.style.display='none'; }
function interactivo(cont, tips, onClick){
  cont.addEventListener('mouseover', function(e){
    var g = e.target.closest ? e.target.closest('[data-t]') : null;
    if(g) tipShow(tips[+g.getAttribute('data-t')], e); });
  cont.addEventListener('mousemove', function(e){
    var g = e.target.closest ? e.target.closest('[data-t]') : null;
    if(g) tipMove(e); else tipHide(); });
  cont.addEventListener('mouseleave', tipHide);
  if(onClick) cont.addEventListener('click', function(e){
    var g = e.target.closest ? e.target.closest('[data-c]') : null;
    if(g){ tipHide(); onClick(g.getAttribute('data-c')); } });
}

/* ============================ tabla reutilizable ============================ */
function estTabla(id, cfg){
  if(!S.tab[id]) S.tab[id] = {c: cfg.orden==null?null:cfg.orden, asc: !!cfg.asc,
                              p:0, per: cfg.per||10, q:''};
  return S.tab[id];
}
function valor(fila, col){ return typeof col.k === 'function' ? col.k(fila) : fila[col.k]; }
function celda(fila, col, mx){
  var v = valor(fila, col);
  if(col.html) return col.html(fila);
  if(col.tipo==='num'){
    var txt = col.fmt ? col.fmt(v,fila) : fnum(v, col.d==null?1:col.d);
    if(col.bar && mx>0 && v>0)
      return '<span class="barcell"><i style="width:'+(v/mx*100).toFixed(1)+
             '%"></i><span>'+txt+'</span></span>';
    return txt;
  }
  return esc(col.fmt ? col.fmt(v,fila) : v);
}
function tabla(cont, id, cfg){
  var st = estTabla(id, cfg);
  var cols = cfg.cols, filas = cfg.filas;
  if(st.q){
    var q = st.q.toLowerCase();
    filas = filas.filter(function(f){
      return cols.some(function(c){
        if(c.tipo==='num') return false;
        return String(c.buscar ? c.buscar(f) : valor(f,c)).toLowerCase().indexOf(q) >= 0; }); });
  }
  if(st.c != null && cols[st.c]){
    var c = cols[st.c];
    filas = filas.slice().sort(function(a,b){
      var x = valor(a,c), y = valor(b,c);
      if(c.tipo==='num'){ x = (x==null||isNaN(x)) ? -Infinity : x;
                          y = (y==null||isNaN(y)) ? -Infinity : y;
                          return st.asc ? x-y : y-x; }
      return st.asc ? String(x).localeCompare(String(y),'es')
                    : String(y).localeCompare(String(x),'es'); });
  }
  var tot = Math.max(1, Math.ceil(filas.length/st.per));
  if(st.p >= tot) st.p = tot-1;
  if(st.p < 0) st.p = 0;
  var vis = filas.slice(st.p*st.per, (st.p+1)*st.per);
  var mx = {};
  cols.forEach(function(c,i){ if(c.bar){ mx[i]=0;
    filas.forEach(function(f){ var v=valor(f,c); if(v>mx[i]) mx[i]=v; }); } });

  var h = ['<div class="tablaBox" data-tid="'+id+'">'];
  if(cfg.barra!==false){
    h.push('<div class="tablaTop"><div class="tt">'+(cfg.titulo||'')+'</div><div class="der">');
    if(cfg.buscar!==false && cfg.filas.length>10)
      h.push('<input type="search" class="tq" placeholder="Buscar" value="'+
             esc(st.q)+'" style="width:150px" aria-label="Buscar en la tabla">');
    if(cfg.csv!==false) h.push('<button class="btn tcsv">Descargar CSV</button>');
    h.push('</div></div>');
  }
  if(!filas.length){
    h.push('<div class="vacio"><b>'+esc(cfg.vacioT||'No hay datos con este filtro')+
      '</b>'+esc(cfg.vacio||'Quita algún filtro de arriba.')+'</div></div>');
    cont.innerHTML = h.join(''); enganchaTabla(cont, id, cfg); return;
  }
  h.push('<div class="scroll"><table><thead><tr>');
  cols.forEach(function(c,i){
    var cls = (c.tipo==='num'?'r ':'') + 's' + (st.c===i ? (st.asc?' asc':' desc') : '');
    h.push('<th class="'+cls+'" data-c="'+i+'" title="'+esc(c.ayuda||('Ordenar por '+c.t))+
      '">'+esc(c.t)+'<span class="ar"></span></th>');
  });
  h.push('</tr></thead><tbody>');
  vis.forEach(function(f){
    h.push('<tr'+(cfg.onRow?' class="click" data-k="'+esc(cfg.rowKey(f))+'"':'')+'>');
    cols.forEach(function(c,i){
      h.push('<td class="'+(c.tipo==='num'?'r ':'')+(c.cls||'')+'">'+celda(f,c,mx[i])+'</td>'); });
    h.push('</tr>');
  });
  h.push('</tbody></table></div>');
  if(filas.length > st.per){
    h.push('<div class="pager"><button class="btn tprev"'+(st.p<=0?' disabled':'')+
      '>Anterior</button><button class="btn tnext"'+(st.p>=tot-1?' disabled':'')+
      '>Siguiente</button><span>'+fnum(st.p*st.per+1,0)+' a '+
      fnum(Math.min((st.p+1)*st.per, filas.length),0)+' de '+fnum(filas.length,0)+
      '</span></div>');
  }
  h.push('</div>');
  cont.innerHTML = h.join('');
  enganchaTabla(cont, id, cfg);
}
function enganchaTabla(cont, id, cfg){
  var st = S.tab[id], box = cont.querySelector('[data-tid]');
  if(!box) return;
  box.querySelectorAll('th.s').forEach(function(th){
    th.addEventListener('click', function(){
      var i = +th.getAttribute('data-c');
      if(st.c === i) st.asc = !st.asc; else { st.c = i; st.asc = cfg.cols[i].tipo!=='num'; }
      st.p = 0; tabla(cont, id, cfg); });
  });
  var q = box.querySelector('.tq');
  if(q) q.addEventListener('input', function(){
    st.q = q.value; st.p = 0; tabla(cont, id, cfg);
    var n = cont.querySelector('.tq'); if(n){ n.focus();
      try{ n.setSelectionRange(n.value.length, n.value.length); }catch(e){} } });
  var pv = box.querySelector('.tprev'), nx = box.querySelector('.tnext');
  if(pv) pv.addEventListener('click', function(){ st.p--; tabla(cont,id,cfg); });
  if(nx) nx.addEventListener('click', function(){ st.p++; tabla(cont,id,cfg); });
  var cs = box.querySelector('.tcsv');
  if(cs) cs.addEventListener('click', function(){ csvTabla(cfg); });
  if(cfg.onRow) box.querySelectorAll('tbody tr.click').forEach(function(tr){
    tr.addEventListener('click', function(){ cfg.onRow(tr.getAttribute('data-k')); }); });
}
function csvTexto(cab, filas){
  var SEP = M.csvSep || ',', DEC = M.csvDec || '.';
  var re = new RegExp('[' + (SEP===';'?';':',') + '"\n\r]');
  var q = function(v){ v = String(v==null?'':v);
    return re.test(v) ? '"'+v.replace(/"/g,'""')+'"' : v; };
  var out = ['sep='+SEP, cab.map(q).join(SEP)];
  filas.forEach(function(f){ out.push(f.map(function(v){
    var t = String(v==null?'':v);
    if(DEC===',') t = t.replace(/^(-?\d+)\.(\d+)$/, '$1,$2');
    return q(t); }).join(SEP)); });
  return '﻿' + out.join('\r\n');
}
function bajaCSV(texto, nombre){
  var a = document.createElement('a');
  a.href = URL.createObjectURL(new Blob([texto], {type:'text/csv;charset=utf-8'}));
  a.download = nombre; document.body.appendChild(a); a.click();
  setTimeout(function(){ URL.revokeObjectURL(a.href); a.remove(); }, 200);
}
function csvTabla(cfg){
  var cols = cfg.cols.filter(function(c){ return c.csv!==false; });
  var filas = cfg.filas.map(function(f){
    return cols.map(function(c){
      if(c.csvk) return c.csvk(f);
      var v = valor(f,c);
      return c.tipo==='num' ? (v==null||isNaN(v)?'':Number(v).toFixed(c.d==null?2:c.d)) : v; }); });
  bajaCSV(csvTexto(cols.map(function(c){return c.t;}), filas), (cfg.archivo||'tabla')+'.csv');
}

/* ============================ helpers svg ============================ */
function ancho(cont, min){ return Math.max(min||280, cont.clientWidth || cont.offsetWidth || 880); }
function esMovil(){ return (window.innerWidth||1024) < 640; }
function vacio(cont, titulo, texto){
  cont.innerHTML = '<div class="vacio"><b>'+esc(titulo)+'</b>'+esc(texto||'')+'</div>';
}
function txtSvg(x,y,t,o){ o=o||{};
  return '<text x="'+x+'" y="'+y+'" font-size="'+(o.s||11)+'" fill="'+(o.c||'#5C7480')+
    '"'+(o.a?' text-anchor="'+o.a+'"':'')+(o.w?' font-weight="'+o.w+'"':'')+
    (o.tr?' transform="'+o.tr+'"':'')+' pointer-events="none">'+esc(t)+'</text>'; }

/* ============================ horas por mes: antes y después ============================ */
function gMeses(cont, rows, leyenda, cmp){
  var todos = MESKEYS.filter(function(k){ return k>=S.m0 && k<=S.m1; });
  if(!rows.length || !todos.length){ vacio(cont,'Sin averías en esta selección',''); leyenda.innerHTML=''; return; }
  var mh = {}, mn = {}, i;
  for(i=0;i<rows.length;i++){ var k = rows[i][F.MK];
    mh[k] = (mh[k]||0) + rows[i][F.H]; mn[k] = (mn[k]||0) + 1; }
  var pos = {}; VM.forEach(function(m,j){ pos[m.k] = j; });
  var lado = function(k){ return pos[k]===undefined ? 'parcial' : (pos[k] < CI ? 'antes' : 'despues'); };
  var COL = {antes:'#B9C8CF', despues:'#10607F', parcial:'#E3EAED'};
  var movil = esMovil();
  var W = ancho(cont), H = movil ? 250 : 300;
  var Mg = {t:26, r:movil?8:96, b:30, l:44}, iw = W-Mg.l-Mg.r, ih = H-Mg.t-Mg.b;
  var mx = 0; todos.forEach(function(k){ if((mh[k]||0) > mx) mx = mh[k]; });
  if(cmp){ mx = Math.max(mx, cmp.antes, cmp.despues); }
  mx = mx*1.12 || 1;
  var paso = iw/todos.length, bw = Math.max(6, Math.min(56, paso*0.62));
  var X = function(j){ return Mg.l + (j+0.5)*paso; };
  var Y = function(v){ return Mg.t + ih - v/mx*ih; };
  var s = ['<svg viewBox="0 0 '+W+' '+H+'" width="100%" height="'+H+'" role="img" '+
           'aria-label="Horas de avería por mes, antes y después">'], tips = [];
  for(var g=0; g<=4; g++){
    var v = mx*g/4;
    s.push('<line x1="'+Mg.l+'" y1="'+Y(v)+'" x2="'+(Mg.l+iw)+'" y2="'+Y(v)+'" stroke="#EEF2F3"/>');
    s.push(txtSvg(Mg.l-7, Y(v)+4, fnum(v,0), {a:'end', s:10.5}));
  }
  s.push(txtSvg(Mg.l, 12, 'Horas de avería en el mes', {s:11.5, c:'#1B3B4A'}));
  todos.forEach(function(k, j){
    var h = mh[k]||0, l = lado(k), c = COV[k];
    var y = Y(h);
    s.push('<rect x="'+(X(j)-bw/2).toFixed(1)+'" y="'+y.toFixed(1)+'" width="'+bw.toFixed(1)+
      '" height="'+Math.max(0.5, Y(0)-y).toFixed(1)+'" fill="'+COL[l]+'" rx="2"/>');
    var etqV = l==='parcial' && !movil ? fnum(h,0)+' en '+fnum(c.d,0)+' '+
      plural(Math.round(c.d),'día','días') : fnum(h,0);
    s.push(txtSvg(X(j), y-5, etqV, {a:'middle', s: movil ? 9.5 : (l==='parcial'?10:11),
      c: l==='despues'?'#0B2430':(l==='parcial'?'#8199A3':'#5C7480'), w: l==='despues'?600:400}));
    s.push(txtSvg(X(j), Mg.t+ih+17, movil ? MESES[k%12] : mesTxt(k), {a:'middle', s:10.5,
      c: l==='parcial'?'#A8B8BF':'#5C7480'}));
    var puedeCorte = pos[k]!==undefined && pos[k] >= 2 && VM.length-pos[k] >= 2;
    tips.push(tipHtml(mayus(mesLargo(k)), [
      ['Horas de avería', fh(h)+' h'],
      ['Averías', fnum(mn[k]||0,0)],
      ['Horas por día', fh(h/Math.max(c.d,1/24))],
      ['Periodo', l==='parcial' ? 'mes incompleto ('+fnum(c.d,1)+' días)' : (l==='antes'?'antes':'después')]],
      puedeCorte ? 'Clic para comparar desde este mes' :
      (l==='parcial' ? 'No entra en la comparación: tiene pocos días de datos.' :
       'Para comparar desde aquí hacen falta 2 meses completos antes y 2 después.')));
    s.push('<rect class="hit" data-t="'+(tips.length-1)+'"'+(puedeCorte?' data-c="'+k+'"':'')+
      ' x="'+(X(j)-paso/2).toFixed(1)+'" y="'+Mg.t+'" width="'+paso.toFixed(1)+'" height="'+ih+
      '" fill="transparent"/>');
  });
  // promedios de cada periodo
  if(cmp && CI>0){
    var ja = todos.indexOf(VM[0].k), jb = todos.indexOf(VM[CI-1].k),
        jc = todos.indexOf(VM[CI].k), jd = todos.indexOf(VM[VM.length-1].k);
    var linea = function(j0, j1, val, col, txt){
      var x0 = X(j0)-paso*0.42, x1 = X(j1)+paso*0.42, yy = Y(val);
      s.push('<line x1="'+x0.toFixed(1)+'" y1="'+yy.toFixed(1)+'" x2="'+x1.toFixed(1)+'" y2="'+
        yy.toFixed(1)+'" stroke="'+col+'" stroke-width="2" stroke-dasharray="6 4"/>');
      return yy;
    };
    var ya = linea(ja, jb, cmp.antes, '#5C7480'), yd = linea(jc, jd, cmp.despues, '#0B2430');
    if(!movil){
      var xr = Mg.l+iw+8;
      var ta = ya, td = yd;
      if(Math.abs(ta-td) < 26){ if(ta<td){ ta -= 13; td += 13; } else { ta += 13; td -= 13; } }
      s.push(txtSvg(xr, ta-2, 'antes', {s:10.5, c:'#5C7480'}));
      s.push(txtSvg(xr, ta+11, fh(cmp.antes)+' h/mes', {s:11.5, c:'#5C7480', w:600}));
      s.push(txtSvg(xr, td-2, 'después', {s:10.5, c:'#0B2430'}));
      s.push(txtSvg(xr, td+11, fh(cmp.despues)+' h/mes', {s:11.5, c:'#0B2430', w:600}));
    }
  }
  s.push('</svg>');
  cont.innerHTML = s.join('');
  interactivo(cont, tips, function(k){ S.corte = +k; render(); });
  var ant = CI>0 && VM.length ? rangoMeses(VM[0].k, VM[CI-1].k) : '',
      des = CI>=0 && VM[CI] ? rangoMeses(VM[CI].k, VM[VM.length-1].k) : '';
  leyenda.innerHTML = (cmp ? '<span><i style="background:'+COL.antes+'"></i>Antes ('+ant+')</span>'+
    '<span><i style="background:'+COL.despues+'"></i>Después ('+des+')</span>' : '')+
    (todos.some(function(k){ return lado(k)==='parcial'; }) ?
      '<span><i style="background:'+COL.parcial+'"></i>Mes incompleto, no se compara</span>' : '')+
    '<span>Línea punteada: promedio por mes de cada periodo. Toca un mes para comparar desde ahí.</span>';
}

/* ============================ barras horizontales ============================ */
function gBarras(cont, grupos, opt){
  opt = opt || {};
  var g = porHoras(grupos), total = suma(g.map(function(x){return x.h;}));
  if(!g.length || total<=0){ vacio(cont,'Sin averías en esta selección',''); return; }
  var TOP = Math.min(opt.top || 8, g.length), movil = esMovil();
  var W = ancho(cont), etq = movil ? 110 : (opt.etq || 220), val = 92;
  var largo = Math.round(etq/6.9), rh = 28, H = TOP*rh + 6;
  var iw = Math.max(40, W - etq - val - 10), mx = g[0].h;
  var s = ['<svg viewBox="0 0 '+W+' '+H+'" width="100%" height="'+H+'">'], tips=[];
  for(var i=0;i<TOP;i++){
    var x = g[i], y = i*rh + 16, w = x.h/mx*iw;
    var filas = [['Horas de avería', fh(x.h)+' h'], ['Parte del total', fpc(x.h/total*100,0)],
                 ['Averías', fnum(x.n,0)], ['Minutos por avería', fmin(x.mttr)]];
    if(x.cmp) filas.push(['Antes → después', antesDespues(x.cmp)+' h/mes']);
    tips.push(tipHtml(x.full, filas, 'Clic para filtrar'));
    s.push('<g class="hit" data-t="'+(tips.length-1)+'" data-c="'+x.key+'">');
    s.push('<rect x="0" y="'+(y-13)+'" width="'+W+'" height="'+(rh-2)+'" fill="transparent"/>');
    s.push(txtSvg(etq-9, y+4, corta(x.label, largo), {s:12.5, c:'#1B3B4A', a:'end', w:500}));
    s.push('<rect x="'+etq+'" y="'+(y-8)+'" width="'+Math.max(1,w).toFixed(1)+
      '" height="17" fill="#10607F" rx="2"/>');
    s.push(txtSvg(etq+w+7, y+4, fh(x.h)+' h', {s:12, c:'#0B2430', w:600}));
    s.push(txtSvg(W-2, y+4, fpc(x.h/total*100,0), {s:11.5, c:'#8199A3', a:'end'}));
    s.push('</g>');
  }
  s.push('</svg>');
  cont.innerHTML = s.join('');
  interactivo(cont, tips, opt.onClick || null);
}

/* ============================ barra de filtros ============================ */
function cuentaH(dimKey){
  var d = DIMS.filter(function(x){return x.k===dimKey;})[0];
  var base = filtra(dimKey), c = {}, i;
  for(i=0;i<base.length;i++) c[base[i][d.f]] = (c[base[i][d.f]]||0) + base[i][F.H];
  return c;
}
function chips(contId, dimKey, todasTxt){
  var c = cuentaH(dimKey), sel = S.sel[dimKey];
  var orden = D.dic[dimKey].map(function(v,i){ return i; })
    .filter(function(i){ return (c[i]||0) > 0 || sel.indexOf(i)>=0; })
    .sort(function(a,b){ return (c[b]||0)-(c[a]||0); });
  var h = '<button class="chip" data-i="all" aria-pressed="'+(!sel.length)+'">'+todasTxt+'</button>' +
    orden.map(function(i){
      return '<button class="chip" aria-pressed="'+(sel.indexOf(i)>=0)+'" data-i="'+i+'" title="'+
        fh(c[i]||0)+' h de avería">'+esc(D.dic[dimKey][i])+'</button>'; }).join('');
  var cont = el(contId); cont.innerHTML = h;
  cont.querySelectorAll('.chip').forEach(function(b){
    b.addEventListener('click', function(){
      var v = b.getAttribute('data-i');
      if(v==='all') S.sel[dimKey] = [];
      else { var i = +v, p = S.sel[dimKey].indexOf(i);
             if(p>=0) S.sel[dimKey].splice(p,1); else S.sel[dimKey].push(i); }
      if(dimKey==='maquina') S.sel.modulo = [];
      render(); }); });
}
function pintaModulos(){
  var c = cuentaH('modulo'), s = el('fmod');
  var idx = D.dic.modulo.map(function(v,i){ return i; })
    .filter(function(i){ return (c[i]||0) > 0 || S.sel.modulo.indexOf(i)>=0; })
    .sort(function(a,b){ return (c[b]||0)-(c[a]||0); });
  s.innerHTML = '<option value="">Todos los módulos</option>' + idx.map(function(i){
    return '<option value="'+i+'">'+esc(corta(modLabel(i),34))+' · '+fh(c[i]||0)+' h</option>'; }).join('');
  s.value = S.sel.modulo.length===1 ? String(S.sel.modulo[0]) : '';
}
function pintaMeses(){
  var s0 = el('fm0'), s1 = el('fm1');
  if(!s0.options.length){
    var op = MESKEYS.map(function(k){ var c = COV[k];
      return '<option value="'+k+'">'+mesCorto(k)+
        (c.d < 0.5*c.dm ? ' ('+fnum(c.d,0)+' '+plural(Math.round(c.d),'día','días')+')' : '')+
        '</option>'; }).join('');
    s0.innerHTML = op; s1.innerHTML = op;
  }
  s0.value = String(S.m0); s1.value = String(S.m1);
}
function pintaCorte(){
  var s = el('fcorte'), op = [];
  VM.forEach(function(m,i){ if(i>=2 && VM.length-i>=2)
    op.push('<option value="'+m.k+'">'+mesLargo(m.k)+'</option>'); });
  s.innerHTML = op.length ? op.join('') : '<option value="">(hacen falta 4 meses completos)</option>';
  s.disabled = !op.length;
  if(op.length && S.corte!=null) s.value = String(S.corte);
}
function pintaPills(){
  var p = [];
  DIMS.forEach(function(d){
    S.sel[d.k].forEach(function(i){
      p.push([d.t+': '+corta(dimTexto(d.k,i),30), d.k+':'+i]); }); });
  if(S.m0!==MESKEYS[0] || S.m1!==MESKEYS[MESKEYS.length-1])
    p.push(['Meses: '+rangoMeses(S.m0, S.m1), 'mes:0']);
  if(S.q) p.push(['Texto: '+S.q, 'q:0']);
  var cont = el('fpills');
  cont.classList.toggle('oculto', !p.length);
  cont.innerHTML = '<span class="fetq">Filtrando</span>' + p.map(function(o){
    return '<button class="pill" data-k="'+esc(o[1])+'" title="Quitar este filtro">'+esc(o[0])+
      ' ✕</button>'; }).join('') + '<button class="btn" data-k="todo">Quitar todos</button>';
  cont.querySelectorAll('[data-k]').forEach(function(b){
    b.addEventListener('click', function(){
      var p2 = b.getAttribute('data-k').split(':'), k = p2[0];
      if(k==='todo'){ limpiar(); return; }
      if(k==='mes'){ S.m0=MESKEYS[0]; S.m1=MESKEYS[MESKEYS.length-1]; }
      else if(k==='q'){ S.q=''; el('fq').value=''; }
      else { var i = S.sel[k].indexOf(+p2[1]); if(i>=0) S.sel[k].splice(i,1); }
      render(); }); });
}
function pintaEstado(rows){
  el('festado').innerHTML = '<b>'+fh(sumaH(rows))+' h</b> en <b>'+fnum(rows.length,0)+'</b> '+
    plural(rows.length,'avería','averías');
}
function limpiar(){
  DIMS.forEach(function(d){ S.sel[d.k] = []; });
  S.q=''; el('fq').value=''; S.m0=MESKEYS[0]; S.m1=MESKEYS[MESKEYS.length-1];
  S.corte = null; render();
}

/* ============================ 1. ¿ESTÁ FUNCIONANDO? ============================ */
function queSeMira(){
  var m = S.sel.maquina, mo = S.sel.modulo, partes = [];
  if(m.length===1 && mo.length===1) return actFull(m[0], mo[0]);
  if(m.length) partes.push(m.map(function(i){ return D.dic.maquina[i]; }).join(', '));
  if(mo.length) partes.push(mo.map(function(i){ return D.dic.modulo[i]; }).join(', '));
  if(S.sel.tipo.length) partes.push('averías '+S.sel.tipo.map(function(i){
    return D.dic.tipo[i].toLowerCase(); }).join(' y '));
  if(S.sel.familia.length) partes.push(S.sel.familia.map(function(i){ return D.dic.familia[i]; }).join(', '));
  if(S.sel.l4.length) partes.push('«'+S.sel.l4.map(function(i){ return D.dic.l4[i]; }).join('», «')+'»');
  return partes.length ? partes.join(' · ') : 'la planta';
}
function veredicto(cont, rows){
  var tot = porMaquina(rows), quien = queSeMira();
  var e = VM.map(function(m){ return m.e; }), mh = {};
  rows.forEach(function(r){ mh[r[F.MK]] = (mh[r[F.MK]]||0) + r[F.H]; });
  var nVM = 0, enVM = {}; VM.forEach(function(m){ enVM[m.k]=1; });
  rows.forEach(function(r){ if(enVM[r[F.MK]]) nVM++; });
  var c = nVM >= M.minVer ? comparar(VM.map(function(m){ return mh[m.k]||0; }), e, CI) : null;
  if(!rows.length){ cont.innerHTML = '<div class="ver"><div class="t">Sin averías</div>'+
    '<div class="s">No hay averías con los filtros actuales. Quita alguno para volver a ver datos.</div></div>';
    return null; }
  if(!c){
    cont.innerHTML = '<div class="ver"><div class="t">Todavía no se puede saber</div><div class="s">'+
      (VM.length < 4 ? 'Para comparar hacen falta al menos 4 meses completos en el rango elegido.'
                     : 'Hay muy pocas averías en '+esc(quien)+' para comparar antes y después.')+
      '</div></div>';
    return null;
  }
  var ant = rangoMeses(VM[0].k, VM[CI-1].k), des = rangoMeses(VM[CI].k, VM[VM.length-1].k);
  var T = {
    MEJORO:'Sí, está funcionando',
    BAJO:'Va mejor, pero todavía no es seguro',
    IGUAL:'No se nota un cambio',
    SUBIO:'Va peor, pero todavía no es seguro',
    EMPEORO:'No, está empeorando'}[c.estado];
  var dif = fh1(Math.abs(c.despues - c.antes));
  var S2 = {
    MEJORO:'Desde '+mesLargo(VM[CI].k)+' '+esc(quien)+' pierde '+dif+' h menos por mes que antes. '+
      'La baja es mayor que lo que varía normalmente de un mes a otro.',
    BAJO:'Desde '+mesLargo(VM[CI].k)+' '+esc(quien)+' pierde '+dif+' h menos por mes, pero los '+
      'meses varían tanto que todavía puede ser casualidad. Conviene esperar uno o dos meses más.',
    IGUAL:'Desde '+mesLargo(VM[CI].k)+' '+esc(quien)+' pierde prácticamente lo mismo por mes que antes.',
    SUBIO:'Desde '+mesLargo(VM[CI].k)+' '+esc(quien)+' pierde '+dif+' h más por mes, pero los '+
      'meses varían tanto que todavía puede ser casualidad. Conviene vigilarlo.',
    EMPEORO:'Desde '+mesLargo(VM[CI].k)+' '+esc(quien)+' pierde '+dif+' h más por mes que antes. '+
      'La subida es mayor que lo que varía normalmente de un mes a otro.'}[c.estado];
  cont.innerHTML = '<div class="ver '+c.estado+'"><div class="t">'+T+'</div><div class="s">'+S2+'</div>'+
    '<div class="cifras">'+
    '<div class="cifra"><div class="v">'+fh1(c.antes)+'<i>h/mes</i></div><div class="k">Antes · '+ant+'</div></div>'+
    '<div class="cifra"><div class="v">'+fh1(c.despues)+'<i>h/mes</i></div><div class="k">Después · '+des+'</div></div>'+
    '<div class="cifra"><div class="v">'+flecha(c)+'</div><div class="k">Cambio</div></div>'+
    '</div></div>';
  return c;
}
function cambios(cont, rows){
  // si se mira más de una máquina se comparan máquinas; si no, sus equipos
  var maqs = porMaquina(rows), porMaq = maqs.length > 1;
  var grupos = (porMaq ? maqs : porActivo(rows)).filter(function(g){ return g.cmp; });
  if(!grupos.length){ cont.innerHTML = ''; return; }
  grupos.forEach(function(g){ g.dif = g.cmp.despues - g.cmp.antes; });
  var baja = grupos.filter(function(g){ return g.dif < 0; }).sort(function(a,b){ return a.dif-b.dif; }).slice(0,4);
  var sube = grupos.filter(function(g){ return g.dif > 0; }).sort(function(a,b){ return b.dif-a.dif; }).slice(0,4);
  var li = function(g){ return '<li data-k="'+(porMaq ? 'm:'+g.maq : 'e:'+g.maq+':'+g.mod)+'">'+
    '<span><span class="nom">'+esc(porMaq ? g.label : g.full)+'</span></span>'+
    '<span class="val">'+antesDespues(g.cmp)+' h/mes '+badgeV(g.cmp)+'</span></li>'; };
  var nada = function(t){ return '<p class="suave" style="margin:4px 0 0;font-size:14px">'+t+'</p>'; };
  var que = porMaq ? 'Las máquinas' : 'Los equipos', ning = porMaq ? 'Ninguna' : 'Ninguno';
  cont.innerHTML =
    '<div><h3>'+que+' que más bajaron</h3>'+(baja.length ? '<ul class="lista">'+
      baja.map(li).join('')+'</ul>' : nada(ning+' bajó.'))+'</div>'+
    '<div><h3>'+que+' que más subieron</h3>'+(sube.length ? '<ul class="lista">'+
      sube.map(li).join('')+'</ul>' : nada(ning+' subió.'))+'</div>';
  cont.querySelectorAll('li[data-k]').forEach(function(x){
    x.addEventListener('click', function(){
      var p = x.getAttribute('data-k').split(':');
      if(p[0]==='m'){ S.sel.maquina = [+p[1]]; S.sel.modulo = []; render(); arriba(); }
      else filtraEquipo(p[1]+':'+p[2]); }); });
}

/* ============================ 2. MÁQUINAS ============================ */
function colCambio(){ return {t:'Antes → después (h/mes)',
  ayuda:'Horas de avería por mes antes y después del mes de comparación',
  k:function(f){ var c=(f._g||f).cmp; return c ? c.despues : -1; }, tipo:'num',
  html:function(f){ return '<span class="num">'+antesDespues((f._g||f).cmp)+'</span>'; },
  csvk:function(f){ return antesDespues((f._g||f).cmp); }}; }
function colVer(){ return {t:'¿Mejoró?',
  ayuda:'«Mejoró» o «Empeoró» solo si el cambio es mayor que la variación normal entre meses',
  k:function(f){ var c=(f._g||f).cmp; return c ? VORD[c.estado] : 9; }, tipo:'num',
  html:function(f){ return badgeV((f._g||f).cmp); },
  csvk:function(f){ var c=(f._g||f).cmp; return c ? VNOM[c.estado] : 'pocos datos'; }}; }
function seccionMaquinas(rows){
  var maqs = porHoras(porMaquina(rows)), tot = sumaH(rows) || 1;
  var filas = maqs.map(function(g){
    var peor=null, pv=-1;
    Object.keys(g.mods).forEach(function(m){ if(g.mods[m]>pv){ pv=g.mods[m]; peor=+m; } });
    return {maq:D.dic.maquina[g.maq], n:g.n, h:g.h, pct:g.h/tot*100, mttr:g.mttr,
            peor:peor, ppeor:g.h>0?pv/g.h*100:0, _g:g}; });
  tabla(el('t-maq'), 'maquinas', {
    titulo:'Toca una máquina para ver solo esa', archivo:'maquinas', per:12, buscar:false,
    cols:[
      {t:'Máquina', k:'maq', cls:'b'},
      {t:'Horas de avería', k:'h', tipo:'num', d:1, bar:true,
       fmt:function(v,f){ return fh(v)+' <span class="suave">· '+fpc(f.pct,0)+'</span>'; }},
      {t:'Averías', k:'n', tipo:'num', d:0},
      {t:'Min. por avería', k:'mttr', tipo:'num', d:0, fmt:function(v){return fmin(v);},
       csvk:function(f){ return (f.mttr*60).toFixed(1); }},
      colCambio(), colVer(),
      {t:'Módulo que más pierde', k:function(f){ return f.peor==null?'':modLabel(f.peor); },
       html:function(f){ return f.peor==null ? '—' : '<span title="'+esc(modLabel(f.peor))+'"><b>'+
         esc(D.dic.modulo[f.peor])+'</b> <span class="suave">'+fpc(f.ppeor,0)+'</span></span>'; }}
    ], filas: filas,
    rowKey:function(f){ return String(f._g.maq); },
    onRow:function(k){ S.sel.maquina = [+k]; S.sel.modulo = []; render(); arriba(); }
  });
  el('leyv').innerHTML = LEYENDA_V;
}

/* ============================ 3. EQUIPOS ============================ */
function seccionEquipos(rows){
  var acts = porActivo(rows);
  comoFalla(acts);
  var rec = CI >= 0 && VM.length > CI;
  acts.forEach(function(g){ var r = masRepetido(g); g.rl4 = r.l4; g.rn = r.n;
    g.ahora = g.cmp ? g.cmp.despues : 0; });
  var filas = acts.filter(function(g){ return D.dic.modulo[g.mod] !== '(sin dato)'; })
    .sort(function(a,b){ return (b.ahora-a.ahora) || (b.h-a.h); });
  el('s3lead').innerHTML = 'Ordenados por lo que pierden <b>ahora</b> (horas por mes '+
    (rec ? 'desde '+mesLargo(VM[CI].k) : 'en el periodo')+'), no por el total, para no poner arriba '+
    'algo que ya se resolvió. Toca un equipo para ver solo ese.';
  tabla(el('t-eq'), 'equipos', {
    titulo:'Un equipo es un módulo de una máquina', archivo:'equipos', per:10,
    cols:[
      {t:'Equipo', k:function(f){ return f.full; },
       html:function(f){ return '<b class="eq">'+esc(f.label)+'</b><span class="sub">'+
         esc(D.modnom[f.mod]||'')+'</span>'; }},
      colCambio(), colVer(),
      {t:'¿Cómo falla?', k:function(f){ return CNOM[f.q]; }, cls:'w2',
       ayuda:'Comparado con el promedio de los equipos',
       html:function(f){ return '<span title="'+esc(CTXT[f.q])+'">'+CNOM[f.q]+'</span>'; }},
      {t:'Lo que más se repite', k:function(f){ return D.dic.l4[f.rl4]; }, cls:'w',
       html:function(f){ return esc(D.dic.l4[f.rl4])+' <span class="suave">('+fnum(f.rn,0)+' de '+
         fnum(f.n,0)+')</span>'+(D.gen[f.rl4]?' <span class="gen">genérico</span>':''); },
       csvk:function(f){ return D.dic.l4[f.rl4]+' ('+f.rn+' de '+f.n+')'; }},
      {t:'Horas totales', k:'h', tipo:'num', d:1, fmt:function(v){return fh(v);},
       ayuda:'Horas de avería en todo el periodo del filtro'}
    ], filas: filas,
    vacioT:'Sin equipos con módulo en esta selección',
    rowKey:function(f){ return f.maq+':'+f.mod; },
    onRow:filtraEquipo
  });
}

/* ============================ 4. QUÉ FALLA ============================ */
function seccionFallas(rows){
  gBarras(el('g-fam'), porDim(rows,'familia'), {top: 8, etq: 240,
    onClick:function(k){ S.sel.familia = [+k]; render(); arriba(); }});
  var g = {}, tot = sumaH(rows) || 1;
  rows.forEach(function(r){ var k = r[F.L4];
    if(!g[k]) g[k] = {l4:k, n:0, h:0, eq:{}};
    g[k].n++; g[k].h += r[F.H]; g[k].eq[r[F.MAQ]+'|'+r[F.MOD]] = 1; });
  var filas = Object.keys(g).map(function(x){ var o = g[x]; o.neq = Object.keys(o.eq).length;
    o.pct = o.h/tot*100; return o; }).sort(function(a,b){ return b.n-a.n; });
  tabla(el('t-l4'), 'detalles', {
    titulo:'Toca un detalle para ver dónde ocurre', archivo:'detalles', per:10,
    cols:[
      {t:'Detalle registrado (L4)', k:function(f){ return D.dic.l4[f.l4]; }, cls:'w',
       html:function(f){ return esc(D.dic.l4[f.l4])+(D.gen[f.l4]?' <span class="gen">genérico</span>':''); }},
      {t:'Veces', k:'n', tipo:'num', d:0, bar:true},
      {t:'Horas', k:'h', tipo:'num', d:1, fmt:function(v){return fh(v);}},
      {t:'% de las horas', k:'pct', tipo:'num', d:0, fmt:function(v){return fpc(v,0);}},
      {t:'En cuántos equipos', k:'neq', tipo:'num', d:0}
    ], filas: filas,
    rowKey:function(f){ return String(f.l4); },
    onRow:function(k){ S.sel.l4 = [+k]; render(); arriba(); }
  });
}
function pintaEventos(rows){
  tabla(el('t-ev'), 'eventos', {
    titulo:'Las averías del filtro actual (los registros de una misma parada ya están unidos)',
    archivo:'averias', per:25, orden:0, asc:false,
    cols:[
      {t:'Inicio', k:function(f){return f[F.T];}, tipo:'num', d:0,
       fmt:function(v){ return fechaHora(v); }, csvk:function(f){ return fechaHora(f[F.T]); }},
      {t:'Máquina', k:function(f){return D.dic.maquina[f[F.MAQ]];}, cls:'b'},
      {t:'Módulo', k:function(f){return modLabel(f[F.MOD]);}},
      {t:'Sistema (L3)', k:function(f){return D.dic.l3[f[F.L3]];}},
      {t:'Detalle (L4)', k:function(f){return D.dic.l4[f[F.L4]];}},
      {t:'Tipo', k:function(f){return D.dic.tipo[f[F.TIP]];}},
      {t:'Turno', k:function(f){return D.dic.turno[f[F.TUR]];}},
      {t:'Minutos', k:function(f){return f[F.H]*60;}, tipo:'num', d:1}
    ], filas: rows, vacioT:'Ninguna avería coincide con el filtro'
  });
}

/* ============================ orquestacion ============================ */
var RENDER_OK = true;
function render(){
  calculaVentana();
  var rows = filtra(null);
  chips('fmaq', 'maquina', 'Todas'); chips('ftipo', 'tipo', 'Todos');
  pintaModulos(); pintaMeses(); pintaCorte(); pintaPills(); pintaEstado(rows);
  try {
    var c = veredicto(el('ver'), rows);
    gMeses(el('g-mes'), rows, el('g-mesleg'), c);
    cambios(el('cambios'), rows);
    seccionMaquinas(rows);
    seccionEquipos(rows);
    seccionFallas(rows);
    if(el('d-ev').open) pintaEventos(rows);
  } catch(err){ if(RENDER_OK){ RENDER_OK = false; console.error('Error al dibujar', err); } }
}
function init(){
  TIP = el('tip');
  D.t0ms = Date.parse(D.t0.replace(' ','T')+'Z');
  D.finms = Date.parse(D.fin.replace(' ','T')+'Z');
  D.cov.forEach(function(c){
    var k = c[0];
    COV[k] = {d:c[1], dm:c[2], a:Math.max(D.t0ms, mesMs(k)), b:Math.min(D.finms, mesMs(k+1))};
    MESKEYS.push(k); });
  D.rows.forEach(function(r){
    r[F.MK] = mesKeyMs(D.t0ms + r[F.T]*3600000);
    if(!COV[r[F.MK]]){ COV[r[F.MK]] = {d:1, dm:30, a:mesMs(r[F.MK]), b:mesMs(r[F.MK]+1)};
                       MESKEYS.push(r[F.MK]); } });
  MESKEYS.sort(function(a,b){ return a-b; });
  S.m0 = MESKEYS[0]; S.m1 = MESKEYS[MESKEYS.length-1]; S.corte = M.corte;

  el('fmod').addEventListener('change', function(){
    S.sel.modulo = this.value==='' ? [] : [+this.value]; render(); });
  el('fm0').addEventListener('change', function(){
    S.m0 = +this.value; if(S.m1 < S.m0) S.m1 = S.m0; render(); });
  el('fm1').addEventListener('change', function(){
    S.m1 = +this.value; if(S.m0 > S.m1) S.m0 = S.m1; render(); });
  el('fcorte').addEventListener('change', function(){
    if(this.value!=='') { S.corte = +this.value; render(); } });
  var tq;
  el('fq').addEventListener('input', function(){
    var v = this.value.toLowerCase(); clearTimeout(tq);
    tq = setTimeout(function(){ S.q = v; render(); }, 250); });
  el('d-ev').addEventListener('toggle', function(){ if(this.open) pintaEventos(filtra(null)); });
  var imp = el('imprimir');
  if(imp) imp.addEventListener('click', function(){ window.print(); });
  document.addEventListener('keydown', function(e){ if(e.key==='Escape') tipHide(); });
  var rt, w0 = window.innerWidth;
  window.addEventListener('resize', function(){
    if(window.innerWidth === w0) return;
    w0 = window.innerWidth; clearTimeout(rt); rt = setTimeout(render, 220); });
  render();
}
if(document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init);
else init();
})();
"""


# =============================================================================
# 7. DOCUMENTO
# =============================================================================
MESES_L = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto",
           "setiembre", "octubre", "noviembre", "diciembre"]
VNOM = {"MEJORO": "mejoró", "BAJO": "bajó (sin confirmar)", "IGUAL": "igual",
        "SUBIO": "subió (sin confirmar)", "EMPEORO": "empeoró"}


def mes_txt(k):
    return f"{MESES_L[k % 12]} de {k // 12}"


def aviso_calidad(dq):
    """Lo que hay que saber antes de buscar causas con el detalle registrado."""
    a = []
    if dq["pct_generico_h"] >= 20:
        a.append(f"En el <b>{dq['pct_generico_h']:.0f} % de las horas</b> el detalle es "
                 f"genérico («Regulación electrónica», «Falla general»…): dice que hubo que "
                 f"intervenir, pero no qué pieza falló.")
    if dq["pct_sin_cod_h"] >= 5:
        txt = (f"El <b>{dq['pct_sin_cod_h']:.0f} % de las horas</b> figura como «Libre» o "
               f"sin clasificar")
        if dq.get("peor_cod") and dq["peor_cod"][1] >= 30:
            txt += (f"; en <b>{dq['peor_cod'][0]}</b> es el {dq['peor_cod'][1]:.0f} % de sus "
                    f"horas")
        a.append(txt + ".")
    if not a:
        return ""
    return ('<div class="nota"><b>Antes de buscar causas con estos detalles:</b> '
            + " ".join(a) + " Para llegar a la causa hay que completar con las órdenes de "
            "trabajo o con quien reparó.</div>")


def notas_html(dq, info, n_av):
    cal = [
        f"<li>{dq['n_reg']:,} registros del archivo se cuentan como {n_av:,} averías: se "
        f"unieron {info['unidos']:,} registros que eran parte de una misma parada (en el "
        f"mismo equipo, el siguiente empezó menos de {UNIR_MINUTOS} minutos después de que "
        f"terminó el anterior; por ejemplo, una parada partida en el cambio de turno).</li>",
        f"<li>{dq['pct_generico_h']:.0f} % de las horas tiene un detalle (L4) genérico, como "
        f"«Regulación electrónica» o «Falla general».</li>",
        f"<li>{dq['pct_sin_cod_h']:.0f} % de las horas figura como «Libre» o sin clasificar"
        + (f"; la máquina con más horas así es {dq['peor_cod'][0]} "
           f"({dq['peor_cod'][1]:.0f} % de sus horas)" if dq.get("peor_cod") else "")
        + ".</li>",
        f"<li>{dq['pct_cortas']:.0f} % de los registros dura menos de 2 minutos y suma solo "
        f"el {dq['pct_cortas_h']:.1f} % de las horas: por eso el reporte se fija en horas y "
        f"no en cantidad de fallas.</li>",
    ]
    if dq.get("unificados"):
        cal.append(f"<li>Se unificaron {dq['unificados']:,} variantes de escritura de "
                   f"Sistema y Detalle (mayúsculas, tildes, prefijos «Elec -» o «Mech -»). "
                   f"Se apaga con UNIFICAR_TEXTOS.</li>")
    for k, dd, dm in dq["cov"]:
        if dd < 0.5 * dm:
            cal.append(f"<li>{mes_txt(k).capitalize()} tiene solo {dd:.1f} días de datos: se "
                       f"ve en el gráfico, pero no entra en la comparación.</li>")
    return f"""
  <ul>
    <li><b>¿Mejoró?</b> Compara las horas de avería por mes antes y después del mes elegido,
      usando solo meses completos (al menos la mitad de los días con datos). Dice
      <b>«Mejoró»</b> o <b>«Empeoró»</b> solo si la diferencia es mayor que lo que varían
      normalmente los meses dentro de cada periodo (razón de tasas cuasi-Poisson, 95 %).
      <b>«Sin confirmar»</b>: va en esa dirección, pero todavía puede ser variación normal.
      <b>«Igual»</b>: el cambio es menor al {CAMBIO_MINIMO} %. Hacen falta 2 meses completos
      antes y 2 después, y al menos {MIN_AVERIAS_VEREDICTO} averías.</li>
    <li>La comparación dice si cambió, no por qué: no descuenta cambios en la producción,
      en los turnos o en la forma de registrar.</li>
    <li><b>¿Cómo falla?</b> Se compara cada equipo con el promedio (método Jack-Knife de
      Knights, 2001): <b>seguido</b> si falla más veces que el promedio de los equipos,
      <b>largo</b> si cada falla dura más que la duración promedio de una falla. «Seguido,
      pero corto» pide evitar que se repita; «pocas veces, pero largo» pide reparar más
      rápido.</li>
    <li><b>Detalle genérico</b>: un detalle que describe la intervención y no la falla
      (regulación, calibración, ajuste electrónico, otros ajustes, falla general, reparación
      mecánica, cambio o reemplazo de componentes, libre).</li>
    {''.join(cal)}
    <li>Con esta data no se puede afirmar la disponibilidad real (falta el tiempo
      programado), el costo de las paradas ni la causa raíz (se registra el síntoma).</li>
  </ul>"""


def construir_html(pl, dq, info):
    m = pl["meta"]
    js = JS.replace("__DATA__", json.dumps(pl, ensure_ascii=False,
                                           separators=(",", ":")).replace("<", "\\u003c"))
    desde, hasta = dq["ini"], dq["fin"] - pd.Timedelta(seconds=1)
    trunc = ("" if not m["truncado"] else
             f'<div class="nota">El reporte contiene las {m["nEmbed"]:,} averías de mayor '
             f'duración de las {m["nFinal"]:,} del periodo. Sube MAX_FILAS_TABLA si necesitas '
             f'todas.</div>')

    return f"""<!DOCTYPE html>
<html lang="es"><head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1,viewport-fit=cover">
<title>Averías · {PLANTA}</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@400;500;600;700&display=swap" rel="stylesheet">
<style>{CSS}</style></head>
<body>

<header class="rail">
  <div class="wrap">
    <div>
      <p class="planta">{PLANTA}</p>
      <h1>{TITULO}</h1>
    </div>
    <div class="railmeta">
      Del {desde:%d/%m/%Y} al {hasta:%d/%m/%Y}<br>
      <b>{m['nFinal']:,}</b> averías · <b>{m['horasTotal']:,.0f} h</b><br>
      Emitido el {m['emitido']}<br>
      <button id="imprimir">Imprimir o guardar en PDF</button>
    </div>
  </div>
</header>

<div class="filtros">
  <div class="wrap">
    <div class="frow">
      <span class="fetq">Máquinas</span>
      <div class="frow" id="fmaq" style="gap:6px"></div>
      <input type="search" id="fq" placeholder="Buscar texto" style="width:150px;margin-left:auto"
             aria-label="Buscar texto libre">
    </div>
    <div class="frow">
      <span class="fetq">Tipo</span>
      <div class="frow" id="ftipo" style="gap:6px"></div>
      <span class="sep"></span>
      <select id="fmod" aria-label="Módulo"></select>
      <span class="sep"></span>
      <span class="suave" style="font-size:13px">Meses</span>
      <select id="fm0" aria-label="Mes inicial"></select>
      <span class="suave" style="font-size:13px">a</span>
      <select id="fm1" aria-label="Mes final"></select>
      <span class="estado" id="festado"></span>
    </div>
    <div class="frow oculto" id="fpills" style="gap:6px"></div>
  </div>
</div>

<main><div class="wrap">
{trunc}
<section class="sec" id="s1">
  <div class="sechead"><span class="n">1</span><h2>¿Está funcionando?</h2>
    <div class="der">Comparar antes y después de <select id="fcorte" aria-label="Mes de comparación"></select></div>
  </div>
  <p class="lead">Elige el mes en que empezó un cambio (un plan, una reparación, un ajuste) y el
  reporte dice si desde entonces se pierden menos horas por averías.</p>
  <div id="ver"></div>
  <div class="chart" id="g-mes"></div>
  <div class="leg" id="g-mesleg"></div>
  <div class="cols c2 cambios" id="cambios" style="margin-top:20px"></div>
</section>

<section class="sec" id="s2">
  <div class="sechead"><span class="n">2</span><h2>¿Dónde se pierden las horas?</h2></div>
  <p class="lead">Una fila por máquina, de la que más horas pierde a la que menos, con su
  propio antes y después.</p>
  <div id="t-maq"></div>
  <p class="leyv" id="leyv"></p>
</section>

<section class="sec" id="s3">
  <div class="sechead"><span class="n">3</span><h2>¿Qué equipos atender primero?</h2></div>
  <p class="lead" id="s3lead"></p>
  <div id="t-eq"></div>
</section>

<section class="sec" id="s4">
  <div class="sechead"><span class="n">4</span><h2>¿Qué falla?</h2></div>
  <p class="lead">Horas por familia de falla y los detalles que más se repiten. Toca una
  barra o una fila para filtrar todo el reporte.</p>
  <h3>Por familia de falla</h3>
  <div class="chart" id="g-fam"></div>
  <h3 style="margin-top:20px">Lo que más se repite</h3>
  <div id="t-l4"></div>
  {aviso_calidad(dq)}
</section>

<details class="bloque" id="d-ev">
  <summary>Ver las averías una por una <span>con los filtros de arriba; se pueden
  descargar en CSV</span></summary>
  <div id="t-ev"></div>
</details>

<details class="bloque">
  <summary>Cómo se calcula y qué tan confiable es el dato</summary>
  {notas_html(dq, info, m['nFinal'])}
</details>

</div></main>

<footer><div class="wrap">{TITULO} · {PLANTA}. Del {desde:%d/%m/%Y} al {hasta:%d/%m/%Y}.
Emitido el {m['emitido']}.</div></footer>

<div class="tip" id="tip"></div>
<script>{js}</script>
</body></html>"""


# =============================================================================
# 8. MAIN
# =============================================================================
def main():
    ruta = buscar_archivo(sys.argv[1] if len(sys.argv) > 1 else ARCHIVO)
    print("=" * 70)
    print("  AVERIAS: DONDE ESTAN LOS PROBLEMAS Y SI LO QUE HACEMOS FUNCIONA  ·  v6")
    print("=" * 70)
    if not ruta or not os.path.exists(ruta):
        print(f"\n  ERROR: no encuentro el archivo:\n  {ruta}")
        print("  Ajusta ARCHIVO al inicio del script o pásalo como argumento.\n")
        return
    print(f"\n  Archivo: {ruta}")

    df = cargar(ruta)
    unificados = df.attrs.get("unificados", 0)
    if FECHA_INICIO:
        df = df[df["Fecha"] >= pd.Timestamp(FECHA_INICIO)]
    if FECHA_FIN:
        df = df[df["Fecha"] < pd.Timestamp(FECHA_FIN) + pd.Timedelta(days=1)]
    if df.empty:
        print("\n  No quedan registros en el rango FECHA_INICIO / FECHA_FIN.\n")
        return

    nombres = cargar_nombres_modulos(df)
    dq = diagnostico(df)
    dq["unificados"] = unificados
    dq["ini"], dq["fin"], dq["cov"] = periodo(df)
    if EXCLUIR_SIN_CODIFICAR:
        df = df[~dq["mask_sin_cod"]]
    av, info = consolidar(df)
    dq["corte"] = corte_inicial(dq["cov"])
    print(f"  {len(df):,} registros → {len(av):,} averías · {av['Horas'].sum():,.0f} h · "
          f"{av['Maquina'].nunique()} máquinas")

    pl = payload(av, nombres, dq)
    salida = SALIDA_HTML or os.path.join(os.path.dirname(os.path.abspath(ruta)),
                                         "reporte_averias.html")
    try:
        with open(salida, "w", encoding="utf-8") as f:
            f.write(construir_html(pl, dq, info))
    except PermissionError:
        base, ext = os.path.splitext(salida)
        salida = f"{base}_{datetime.now():%Y%m%d_%H%M%S}{ext}"
        with open(salida, "w", encoding="utf-8") as f:
            f.write(construir_html(pl, dq, info))
    print(f"  Reporte: {salida} ({os.path.getsize(salida) / 1024:,.0f} KB)")

    planta, maqs = resumen_comparacion(av, dq["cov"], dq["corte"])
    print("\n" + "=" * 70)
    if planta:
        print(f"  ¿Funciona? Desde {mes_txt(dq['corte'])}: {planta['antes']:,.0f} → "
              f"{planta['despues']:,.0f} h de avería por mes "
              f"({planta['cambio']:+.0f} %) · {VNOM[planta['estado']]}")
        for mq, c in sorted(maqs.items(), key=lambda kv: -(kv[1] or {}).get("despues", 0)):
            if c:
                print(f"    {mq:<6} {c['antes']:6.1f} → {c['despues']:6.1f} h/mes "
                      f"({c['cambio']:+4.0f} %) · {VNOM[c['estado']]}")
    else:
        print("  ¿Funciona? Hacen falta al menos 4 meses completos para comparar.")
    print("=" * 70)


if __name__ == "__main__":
    main()
