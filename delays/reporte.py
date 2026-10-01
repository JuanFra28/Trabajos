"""
===============================================================================
 ANALISIS DE CONFIABILIDAD DE AVERIAS  ·  v5
===============================================================================
 Reporte HTML autocontenido e interactivo, pensado para presentar a
 Mantenimiento y a áreas que no son de mantenimiento.

 Novedades v5:
   · Más corto: 5 pestañas (Resumen, Máquinas, Prioridades, Fallas, Notas)
     en lugar de 9. El Resumen se lee en un minuto.
   · "Dónde actuar primero": los equipos que más horas costaron en los
     últimos meses, con el síntoma que más se repite y un siguiente paso.
   · Las paradas partidas en el cambio de turno se unen: antes una parada de
     18 h contaba como 4 fallas.
   · Tendencia más sobria: compara las fallas por mes de la segunda mitad del
     periodo con las de la primera, y solo marca "empeora" o "mejora" si la
     diferencia es mayor que la variación normal de un mes a otro. Crow-AMSAA
     tomaba las rachas de un mes malo como tendencia y marcaba equipos de más.
   · Frecuencia contra duración (Jack-Knife) con los límites de Knights y la
     duración en minutos.
   · Se quitaron Weibull por módulo y la "disponibilidad": con estos datos no
     ayudaban a decidir (la pestaña Notas explica por qué).

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
HOJA = "Datos"                 # si no existe, toma la primera hoja
SALIDA_HTML = None             # None: reporte_confiabilidad.html junto al Excel

PLANTA = "Planta Santa Clara"
TITULO = "Confiabilidad de averías"
SUBTITULO = "Dónde se pierden las horas de producción y por dónde empezar"

# --- Parámetros de análisis
UNIR_MINUTOS = 15            # une registros del mismo equipo cuando entre el fin
                             # de uno y el inicio del siguiente hay menos de N min
FECHA_INICIO = None          # "2026-02-01" para recortar el inicio
FECHA_FIN = None             # "2026-08-31" para recortar el final
EXCLUIR_SIN_CODIFICAR = False
MIN_FALLAS_GRAFICO = 8       # mínimo de fallas para aparecer en Frecuencia vs duración
MIN_FALLAS_TENDENCIA = 15    # mínimo de fallas para opinar sobre la tendencia
MIN_MESES_TENDENCIA = 5      # mínimo de meses completos para la tendencia
MESES_RECIENTES = 3          # "últimos N meses" con que se ordena "Dónde actuar"
TOP_ACCIONES = 5             # equipos en "Dónde actuar primero"
MAX_FILAS_TABLA = 30000      # averías embebidas en el reporte

# --- Textos de falla (Sistema L3 y Detalle L4).
#     True: cuenta como uno solo "Regulacion Electronica", "Regulación electrónica"
#     y "Elec - Regulacion Electronica". False: los deja tal cual vienen.
UNIFICAR_TEXTOS = True

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
# 1. ESTADISTICA  (mismas fórmulas que usa el navegador)
# =============================================================================
# Cuantil 97.5 % de la t de Student, por grados de libertad
T975 = [None, 12.706, 4.303, 3.182, 2.776, 2.571, 2.447, 2.365, 2.306, 2.262, 2.228,
        2.201, 2.179, 2.160, 2.145, 2.131, 2.120, 2.110, 2.101, 2.093, 2.086,
        2.080, 2.074, 2.069, 2.064, 2.060, 2.056, 2.052, 2.048, 2.045, 2.042]


def t975(gl):
    return T975[gl] if gl < len(T975) else 1.96 + 2.4 / gl


def tendencia(y, e):
    """¿Fallan más (o menos) en la segunda mitad del periodo que en la primera?

    y: fallas de cada mes completo; e: días observados de ese mes.
    Compara el ritmo de fallas de la segunda mitad de los meses con el de la
    primera (prueba de razón de tasas) y mide el ruido real con la variación
    de mes a mes dentro de cada mitad (dispersión de Pearson, cuasi-Poisson).
    Así un mes malo aislado no alcanza para decir que el equipo empeora.
    Devuelve SUBE / BAJA cuando la diferencia supera el límite de 95 %, IGUAL
    si no, y None si no alcanzan los datos."""
    y = np.asarray(y, float)
    e = np.asarray(e, float)
    k = len(y)
    if k < MIN_MESES_TENDENCIA or y.sum() < MIN_FALLAS_TENDENCIA:
        return None
    h = k // 2
    y1, y2, e1, e2 = y[:h].sum(), y[h:].sum(), e[:h].sum(), e[h:].sum()
    if e1 <= 0 or e2 <= 0:
        return None
    r1, r2 = y1 / e1, y2 / e2
    mu = np.r_[r1 * e[:h], r2 * e[h:]]
    ok = mu > 0
    phi = max(1.0, float(((y[ok] - mu[ok]) ** 2 / mu[ok]).sum()) / (k - 2))
    n, p0 = y1 + y2, e2 / (e1 + e2)
    z = (y2 - n * p0) / math.sqrt(phi * n * p0 * (1 - p0))
    lim = t975(k - 2)
    return dict(z=float(z), phi=phi, ini=r1 * 30.44, fin=r2 * 30.44,
                estado="SUBE" if z > lim else ("BAJA" if z < -lim else "IGUAL"))


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


def exportar_pendientes(df, nombres, carpeta):
    g = (df.groupby("Modulo").agg(eventos=("Horas", "size"),
                                  horas=("Horas", "sum")).reset_index())
    g["horas"] = g["horas"].round(1)
    g = g[~g["Modulo"].str.upper().isin(nombres.keys()) & (g["Modulo"] != SIN)]
    if g.empty:
        return 0
    out = g.sort_values("horas", ascending=False).rename(columns={"Modulo": "codigo"})
    out.insert(1, "nombre", "")
    out.to_csv(os.path.join(carpeta, "modulos_pendientes.csv"), index=False,
               encoding="utf-8-sig")
    return len(g)


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
    d["n_sin_cod"] = int(mask.sum())
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

    mens = df.set_index("Fecha").resample("MS").size()
    d["sugerido_inicio"] = None
    if len(mens) >= 4:
        pobres = mens[mens < mens.median() * .25]
        if len(pobres) and pobres.index[0] == mens.index[0]:
            arranque = pobres[pobres.index <= mens.index[len(mens) // 2]]
            if len(arranque):
                d["sugerido_inicio"] = arranque.index[-1] + pd.offsets.MonthBegin(1)
    return d


def periodo(df):
    """Inicio y fin observados, y días cubiertos de cada mes."""
    f0 = df["Fecha"].min()
    f1 = (df["Fecha"] + pd.to_timedelta(df["Horas"], unit="h")).max()
    ini = f0.to_period("M").to_timestamp() if f0.day <= 5 else f0.normalize()
    fin = f1
    # el recorte manual nunca extiende el periodo más allá de los datos: un mes
    # sin datos contaría como un mes sin fallas y falsearía la tendencia
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
    """Meses con al menos la mitad de los días observados (para tendencias)."""
    return [(k, dd) for k, dd, dm in cov if dd >= 0.5 * dm]


def analisis_activos(av, cov):
    """Lo mismo que calcula el navegador sin filtros: horas, grupo de
    frecuencia/duración (límites de Knights), tendencia y horas recientes."""
    g = av.groupby("activo").agg(n=("Horas", "size"), h=("Horas", "sum"),
                                 mod=("Modulo", "first"))
    g["mttr"] = g["h"] / g["n"]
    ln, lm = g["n"].sum() / len(g), g["h"].sum() / g["n"].sum()
    g["q"] = np.where((g["n"] > ln) & (g["mttr"] > lm), "CRITICO",
             np.where(g["n"] > ln, "CRONICO",
             np.where(g["mttr"] > lm, "AGUDO", "LEVE")))
    vm = meses_completos(cov)
    ks, e = [k for k, _ in vm], np.array([dd for _, dd in vm], float)
    mk = av["Fecha"].dt.year * 12 + av["Fecha"].dt.month - 1
    tab = pd.crosstab(av["activo"], mk).reindex(index=g.index, columns=ks, fill_value=0)
    g["tend"] = [tendencia(tab.loc[a].to_numpy(float), e) for a in g.index]
    rec = ks[-MESES_RECIENTES:] if len(ks) > MESES_RECIENTES else []
    g["h_rec"] = (av[mk.isin(rec)].groupby("activo")["Horas"].sum()
                  .reindex(g.index, fill_value=0) if rec else 0.0)
    return g, rec


# =============================================================================
# 4. PAYLOAD PARA EL NAVEGADOR
# =============================================================================
DIMS = ["Maquina", "Unidad", "Modulo", "Familia", "Tipo", "Area", "Turno", "L3", "L4"]


def payload(av, nombres, dq, info, limite=MAX_FILAS_TABLA):
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

    sug = dq.get("sugerido_inicio")
    return dict(
        dic={k.lower(): v for k, v in dic.items()},
        modnom=[nombres.get(m.upper(), "") for m in dic["Modulo"]],
        gen=gen, rows=filas, cov=dq["cov"],
        t0=dq["ini"].strftime("%Y-%m-%d %H:%M:%S"),
        fin=dq["fin"].strftime("%Y-%m-%d %H:%M:%S"),
        meta=dict(
            planta=PLANTA, titulo=TITULO, subtitulo=SUBTITULO,
            emitido=datetime.now().strftime("%d/%m/%Y"),
            nReg=int(dq["n_reg"]), nFinal=int(len(av)), nEmbed=len(d),
            truncado=bool(truncado), horasTotal=float(av["Horas"].sum()),
            ventana=UNIR_MINUTOS, nUnidos=int(info["unidos"]),
            nContinuas=int(info["continuas"]),
            pctSinCodH=round(dq["pct_sin_cod_h"], 1), nSinCod=int(dq["n_sin_cod"]),
            horasSinCod=round(dq["horas_sin_cod"], 1),
            peorCod=(list(dq["peor_cod"]) if dq.get("peor_cod") else None),
            pctGenericoH=round(dq["pct_generico_h"], 1),
            pctCortas=round(dq["pct_cortas"], 1), pctCortasH=round(dq["pct_cortas_h"], 1),
            unificados=int(dq.get("unificados", 0)),
            sugerido=(sug.strftime("%Y-%m") if sug is not None else None),
            cobertura=round(float(dq["cobertura"]), 0),
            minJK=MIN_FALLAS_GRAFICO, minTend=MIN_FALLAS_TENDENCIA,
            minMeses=MIN_MESES_TENDENCIA, recientes=MESES_RECIENTES,
            topAcc=TOP_ACCIONES, idioma=IDIOMA,
            csvSep=(";" if CSV_ESTILO == "coma" else ","),
            csvDec=("," if CSV_ESTILO == "coma" else "."),
        ))


# =============================================================================
# 5. ESTILOS
# =============================================================================
CSS = r"""
:root{
  --tinta:#0B2430; --tinta2:#1B3B4A; --tinta3:#34586A;
  --gris:#5C7480; --gris2:#8199A3;
  --linea:#D5DEE0; --linea2:#E9EEEF;
  --fondo:#EDF1F1; --panel:#FFFFFF; --panel2:#F6F9F9;
  --rojo:#B5352B; --rojoBg:#FBEDEB; --rojoBd:#EED2CE;
  --ambar:#9C6B00; --ambarF:#E0A200; --ambarBg:#FCF5E4; --ambarBd:#EDDCB6;
  --verde:#17714B; --verdeBg:#E9F4EF; --verdeBd:#C6E1D4;
  --azul:#10607F; --azul2:#2E89AC; --azul3:#93C6D8; --azulBg:#EAF2F6;
  --r:4px;
  --sombra:0 10px 34px rgba(11,36,48,.14);
  --sans:"IBM Plex Sans","Segoe UI",system-ui,-apple-system,Roboto,Arial,sans-serif;
  --mono:"IBM Plex Mono","Cascadia Mono",Consolas,"SF Mono",Menlo,monospace;
}
*,*::before,*::after{box-sizing:border-box}
html{-webkit-text-size-adjust:100%}
body{margin:0;background:var(--fondo);color:var(--tinta2);font-family:var(--sans);
 font-size:15.5px;line-height:1.55;font-feature-settings:"tnum" 1,"lnum" 1}
h1,h2,h3,h4{margin:0;font-weight:600;color:var(--tinta);letter-spacing:-.011em}
p{margin:0 0 10px}
a{color:var(--azul)}
.wrap{max-width:1240px;margin:0 auto;padding:0 26px}
.oculto{display:none !important}
.suave{color:var(--gris)}

/* ---------- cabecera ---------- */
.rail{background:var(--tinta);color:#fff;position:relative}
.rail::after{content:"";position:absolute;left:0;right:0;bottom:0;height:2px;
 background:linear-gradient(90deg,var(--rojo) 0 33%,var(--ambarF) 33% 66%,var(--verde) 66%)}
.rail .wrap{display:flex;align-items:flex-end;justify-content:space-between;
 gap:26px;padding-top:24px;padding-bottom:18px;flex-wrap:wrap}
.marca h1{color:#fff;font-size:28px;line-height:1.12;font-weight:600}
.marca .planta{color:var(--azul3);font-size:14px;font-weight:500;margin:0 0 6px}
.marca .sub{color:#A9C4CE;font-size:14.5px;margin:7px 0 0;max-width:62ch}
.railmeta{color:#A9C4CE;font-size:13.5px;text-align:right;line-height:1.7}
.railmeta b{color:#fff;font-weight:600}
.railmeta .imp{background:none;border:1px solid rgba(255,255,255,.28);color:#D7E6EC;
 padding:6px 13px;border-radius:var(--r);font-size:13px;cursor:pointer;margin-top:8px;
 font-family:inherit}
.railmeta .imp:hover{background:rgba(255,255,255,.1)}

/* ---------- pestañas ---------- */
.tabs{background:var(--tinta2);position:sticky;top:0;z-index:40;
 box-shadow:0 1px 0 rgba(0,0,0,.2)}
.tabs .wrap{display:flex;gap:2px;overflow-x:auto;scrollbar-width:none}
.tabs .wrap::-webkit-scrollbar{display:none}
.tab{appearance:none;background:none;border:0;border-bottom:3px solid transparent;
 color:#A9C4CE;padding:13px 18px;font-size:14.8px;font-family:inherit;font-weight:500;
 cursor:pointer;white-space:nowrap;line-height:1.2}
.tab:hover{color:#fff;background:rgba(255,255,255,.06)}
.tab[aria-selected="true"]{color:#fff;border-bottom-color:var(--ambarF);font-weight:600}
.tab:focus-visible{outline:2px solid var(--ambarF);outline-offset:-3px}
.tabs .filtroact{margin-left:auto;align-self:center;color:var(--ambarF);font-size:12.8px;
 white-space:nowrap;padding-left:12px}

/* ---------- filtros ---------- */
.filtros{background:var(--panel);border-bottom:1px solid var(--linea)}
.filtros .wrap{padding-top:13px;padding-bottom:13px}
.frow{display:flex;gap:9px;align-items:center;flex-wrap:wrap}
.frow+.frow{margin-top:10px}
.fetq{font-size:12.5px;color:var(--gris);font-weight:600;min-width:62px}
.chip{appearance:none;font-family:inherit;background:var(--panel);
 border:1px solid var(--linea);color:var(--tinta2);padding:6px 12px;border-radius:var(--r);
 font-size:13.8px;font-weight:500;cursor:pointer;line-height:1.3}
.chip:hover{border-color:var(--azul);background:var(--azulBg)}
.chip[aria-pressed="true"]{background:var(--azul);border-color:var(--azul);color:#fff}
.chip.na{opacity:.34;cursor:not-allowed}
.chip .n{opacity:.62;font-size:12.4px;margin-left:5px;font-variant-numeric:tabular-nums}
.chip:focus-visible,.msel>button:focus-visible{outline:2px solid var(--tinta);outline-offset:1px}
.msel{position:relative}
.msel>button{appearance:none;font-family:inherit;background:var(--panel);
 border:1px solid var(--linea);color:var(--tinta2);padding:7px 30px 7px 12px;
 border-radius:var(--r);font-size:13.8px;cursor:pointer;min-width:128px;text-align:left;
 position:relative;line-height:1.3}
.msel>button::after{content:"";position:absolute;right:11px;top:50%;margin-top:-2px;
 border:4px solid transparent;border-top-color:var(--gris)}
.msel>button:hover{border-color:var(--azul)}
.msel>button.act{border-color:var(--azul);background:var(--azulBg);color:var(--azul);
 font-weight:600}
.msel .pop{position:absolute;z-index:60;top:calc(100% + 5px);left:0;width:290px;
 background:var(--panel);border:1px solid var(--linea);border-radius:var(--r);
 box-shadow:var(--sombra);padding:10px;display:none}
.msel.abierto .pop{display:block}
.msel .pop input[type=search]{width:100%;margin-bottom:8px}
.msel .lista{max-height:272px;overflow:auto;margin:0 -4px}
.msel .lista label{display:flex;gap:9px;align-items:center;padding:5px 8px;
 border-radius:var(--r);font-size:13.6px;cursor:pointer}
.msel .lista label:hover{background:var(--panel2)}
.msel .lista label.na{opacity:.4}
.msel .lista .c{margin-left:auto;color:var(--gris2);font-size:12.4px}
.msel .pie{display:flex;gap:8px;border-top:1px solid var(--linea2);margin-top:8px;
 padding-top:8px}
.msel .pie button{flex:1;background:var(--panel2);border:1px solid var(--linea);
 color:var(--tinta2);font-size:13px;padding:6px;border-radius:var(--r);cursor:pointer;
 font-family:inherit}
.msel .pie button:hover{background:var(--azulBg);border-color:var(--azul)}
input[type=search],select{font-family:inherit;font-size:13.8px;
 padding:7px 10px;border:1px solid var(--linea);border-radius:var(--r);
 background:var(--panel);color:var(--tinta2)}
input:focus,select:focus{outline:2px solid var(--azul);outline-offset:-1px}
.btn{appearance:none;font-family:inherit;background:var(--azul);border:1px solid var(--azul);
 color:#fff;padding:7px 14px;border-radius:var(--r);font-size:13.8px;font-weight:500;
 cursor:pointer}
.btn:hover{background:#0C5068}
.btn.g{background:var(--panel);color:var(--azul)}
.btn.g:hover{background:var(--azulBg)}
.btn:focus-visible{outline:2px solid var(--tinta);outline-offset:2px}
.btn:disabled{opacity:.4;cursor:default}
.estado{font-size:13.2px;color:var(--gris);margin-left:auto;text-align:right}
.estado b{color:var(--tinta);font-weight:600}

/* ---------- vistas ---------- */
main{padding:24px 0 56px}
.vista{display:none}
.vista.on{display:block}
.vhead{margin-bottom:16px}
.vhead h2{font-size:23px}
.vhead p{color:var(--gris);font-size:15px;max-width:76ch;margin:5px 0 0}
.card{background:var(--panel);border:1px solid var(--linea);border-radius:var(--r);
 padding:20px 22px;margin-bottom:18px}
.card>h3{font-size:17px;margin-bottom:3px}
.card>.lead{color:var(--gris);font-size:14.2px;margin:0 0 14px;max-width:78ch}
.cols{display:grid;gap:18px}
.cols.c2{grid-template-columns:minmax(0,1fr) minmax(0,1fr)}
.cols.c4{grid-template-columns:repeat(4,minmax(0,1fr))}
.cols>.card{margin-bottom:18px}

/* ---------- cifras ---------- */
.kpis{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));
 gap:1px;background:var(--linea);border:1px solid var(--linea);border-radius:var(--r);
 overflow:hidden;margin-bottom:18px}
.kpi{background:var(--panel);padding:15px 18px}
.kpi .v{font-size:28px;font-weight:600;color:var(--tinta);line-height:1.05;
 display:block;font-variant-numeric:tabular-nums}
.kpi .v i{font-style:normal;font-size:14px;color:var(--gris);font-weight:500;margin-left:4px}
.kpi .k{font-size:13px;color:var(--gris);margin-top:6px}

/* ---------- titular ---------- */
.titular{background:var(--tinta);color:#fff;border-radius:var(--r);padding:22px 26px;
 margin-bottom:18px}
.titular p{font-size:18.5px;line-height:1.5;color:#E6F0F3;margin:0;max-width:74ch}
.titular p+p{margin-top:10px;font-size:15.5px;color:#BCD3DB}
.titular b{color:#fff;font-weight:600}

/* ---------- dónde actuar ---------- */
.accion{border:1px solid var(--linea);border-left:4px solid var(--azul);background:var(--panel);
 padding:14px 18px;border-radius:0 var(--r) var(--r) 0;margin-bottom:10px;cursor:pointer}
.accion:hover{background:#FAFCFC;border-color:var(--azul3)}
.accion.CRITICO{border-left-color:var(--rojo)}
.accion.CRONICO{border-left-color:#D07A1E}
.accion.AGUDO{border-left-color:var(--azul)}
.accion.LEVE{border-left-color:var(--gris2)}
.accion .cab{display:flex;gap:10px;align-items:baseline;flex-wrap:wrap}
.accion .pos{font-size:13px;color:var(--gris2);font-weight:600;min-width:16px}
.accion .t{font-weight:600;color:var(--tinta);font-size:16.5px}
.accion .tags{display:inline-flex;gap:6px;flex-wrap:wrap}
.accion .m{color:var(--gris);font-size:13.6px;margin:4px 0 0 26px}
.accion .m b{color:var(--tinta2);font-weight:600}
.accion .p{font-size:14.4px;margin:7px 0 0 26px;max-width:92ch}
.accion .p b{color:var(--tinta)}

/* ---------- grupos de prioridad ---------- */
.qcard{border:1px solid var(--linea);border-top:4px solid var(--gris2);border-radius:var(--r);
 padding:13px 15px;background:var(--panel)}
.qcard .v{font-size:23px;font-weight:600;color:var(--tinta);line-height:1.1}
.qcard .v i{font-style:normal;font-size:13.4px;color:var(--gris);font-weight:500;margin-left:4px}
.qcard .k{font-size:13.4px;color:var(--tinta2);font-weight:600;margin-top:4px}
.qcard .d{font-size:13px;color:var(--gris);margin-top:2px}
.lista2{margin:0;padding:0;list-style:none}
.lista2 li{display:flex;gap:10px;justify-content:space-between;align-items:baseline;
 padding:7px 0;border-bottom:1px solid var(--linea2);font-size:14px;cursor:pointer}
.lista2 li:last-child{border-bottom:0}
.lista2 li:hover{color:var(--azul)}
.lista2 .n{color:var(--gris);font-size:13px;white-space:nowrap;text-align:right}
.lista2 .n b{color:var(--tinta)}

/* ---------- gráficos ---------- */
.chart{position:relative;width:100%}
.chart svg{display:block;width:100%;height:auto;overflow:visible}
.chart .hit{cursor:pointer}
.tip{position:fixed;z-index:200;display:none;pointer-events:none;width:270px;
 background:var(--tinta);color:#DCE9ED;border-radius:var(--r);padding:11px 13px;
 font-size:12.9px;line-height:1.45;box-shadow:var(--sombra)}
.tip .t{font-weight:600;color:#fff;font-size:13.4px;margin-bottom:7px;padding-bottom:6px;
 border-bottom:1px solid rgba(255,255,255,.17)}
.tip .r{display:flex;justify-content:space-between;gap:12px;padding:2px 0}
.tip .r span{color:#9FBAC4}
.tip .r b{color:#fff;font-weight:600;font-variant-numeric:tabular-nums}
.tip .q{color:#9FBAC4;font-size:11.9px;margin-top:7px;padding-top:7px;
 border-top:1px solid rgba(255,255,255,.15)}
.tip i.dot{width:8px;height:8px;border-radius:50%;display:inline-block;margin-right:6px}
.leg{display:flex;flex-wrap:wrap;gap:7px;align-items:center;margin-top:12px}
.leg .item{display:inline-flex;align-items:center;gap:6px;font-size:12.8px;color:var(--tinta2)}
.leg .item i{width:11px;height:11px;border-radius:2px;display:inline-block}
.leg button.item{appearance:none;font-family:inherit;background:var(--panel);
 border:1px solid var(--linea);padding:5px 10px;border-radius:var(--r);cursor:pointer}
.leg button.item:hover{border-color:var(--gris2)}
.leg button.item.off{opacity:.42;text-decoration:line-through}
.leg .hint{color:var(--gris2);font-size:12.5px}
.leg b{font-variant-numeric:tabular-nums;color:var(--gris)}

/* ---------- cómo se lee ---------- */
details.lee{border:1px solid var(--linea);border-left:3px solid var(--azul2);
 border-radius:0 var(--r) var(--r) 0;background:var(--panel2);margin-top:14px}
details.lee>summary{cursor:pointer;padding:9px 15px;font-size:13.8px;font-weight:600;
 color:var(--azul);list-style:none}
details.lee>summary::-webkit-details-marker{display:none}
details.lee>summary::before{content:"?";display:inline-grid;place-items:center;
 width:17px;height:17px;border-radius:50%;background:var(--azul);color:#fff;
 font-size:11.5px;margin-right:8px;vertical-align:-3px;font-weight:700}
details.lee[open]>summary{border-bottom:1px solid var(--linea)}
details.lee .cuerpo{padding:13px 16px;font-size:14.2px;color:var(--tinta2)}
details.lee .cuerpo p{max-width:78ch}
details.lee .cuerpo p:last-child{margin-bottom:0}
details.bloque>summary{cursor:pointer;font-size:17px;font-weight:600;color:var(--tinta);
 list-style:none}
details.bloque>summary::-webkit-details-marker{display:none}
details.bloque>summary::before{content:"\25B8";display:inline-block;margin-right:8px;
 color:var(--azul);transition:transform .15s}
details.bloque[open]>summary::before{transform:rotate(90deg)}
details.bloque>summary span{font-size:14px;font-weight:400;color:var(--gris);margin-left:6px}
details.bloque[open]>summary{margin-bottom:14px}

/* ---------- avisos ---------- */
.nota{border:1px solid var(--ambarBd);border-left:3px solid var(--ambarF);
 background:var(--ambarBg);padding:13px 17px;border-radius:0 var(--r) var(--r) 0;
 margin-bottom:14px;font-size:14.3px}
.nota.ok{border-color:var(--verdeBd);border-left-color:var(--verde);background:var(--verdeBg)}
.nota b{color:var(--tinta)}
.nota p:last-child,.nota ul:last-child{margin-bottom:0}
.nota ul{margin:6px 0 0;padding-left:19px}
.nota li{margin-bottom:4px}

/* ---------- tablas ---------- */
.tablaBox{border:1px solid var(--linea);border-radius:var(--r);overflow:hidden;
 background:var(--panel)}
.tablaTop{display:flex;gap:10px;align-items:center;padding:10px 12px;
 border-bottom:1px solid var(--linea);background:var(--panel2);flex-wrap:wrap}
.tablaTop .tt{font-size:13.6px;color:var(--gris)}
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
tbody tr:nth-child(even) td{background:var(--panel2)}
tbody tr.click{cursor:pointer}
tbody tr.click:hover td{background:var(--azulBg)}
td.r,th.r{text-align:right}
td.b{font-weight:600;color:var(--tinta)}
td.w{white-space:normal;min-width:200px}
.barcell{position:relative;display:block;min-width:92px;text-align:right}
.barcell i{position:absolute;left:0;top:50%;transform:translateY(-50%);height:16px;
 background:var(--azul3);border-radius:1px;opacity:.45}
.barcell span{position:relative}
.pager{display:flex;gap:9px;align-items:center;padding:9px 12px;font-size:13.3px;
 color:var(--gris);border-top:1px solid var(--linea);flex-wrap:wrap;background:var(--panel)}
.pager .btn{padding:5px 11px;font-size:13.2px}
.vacio{padding:30px 20px;text-align:center;color:var(--gris);font-size:14.4px}
.vacio b{display:block;color:var(--tinta);font-size:15.4px;margin-bottom:5px}

/* ---------- etiquetas ---------- */
.badge{display:inline-block;padding:2px 9px;border-radius:var(--r);font-size:12.2px;
 font-weight:600;border:1px solid;white-space:nowrap;line-height:1.5}
.b-det{color:var(--rojo);background:var(--rojoBg);border-color:var(--rojoBd)}
.b-mej{color:var(--verde);background:var(--verdeBg);border-color:var(--verdeBd)}
.b-nd{color:var(--gris);background:var(--panel2);border-color:var(--linea)}
.b-cri{color:#fff;background:var(--rojo);border-color:var(--rojo)}
.b-cro{color:#8A3E00;background:#FCF0E6;border-color:#EFD6C1}
.b-agu{color:var(--azul);background:var(--azulBg);border-color:#C6DEE8}
.b-lev{color:var(--gris);background:var(--panel2);border-color:var(--linea)}
.gen{color:var(--ambar);font-size:12.4px;white-space:nowrap}

/* ---------- notas ---------- */
.defs{display:grid;grid-template-columns:repeat(auto-fit,minmax(300px,1fr));gap:14px 26px}
.defs div h4{font-size:15px;margin-bottom:3px}
.defs div p{font-size:14.2px;margin:0;color:var(--tinta2)}
.card ul.simple{margin:0;padding-left:19px;font-size:14.3px;max-width:86ch}
.card ul.simple li{margin-bottom:6px}

footer{background:var(--tinta);color:#9FBAC4;padding:20px 0;font-size:13px}
footer .wrap{display:flex;justify-content:space-between;gap:20px;flex-wrap:wrap}
footer b{color:#D7E6EC;font-weight:600}
footer p{max-width:70ch;margin:0}

@media (max-width:1000px){
  .cols.c2{grid-template-columns:minmax(0,1fr)}
  .cols.c4{grid-template-columns:repeat(2,minmax(0,1fr))}
  .kpis{grid-template-columns:repeat(2,minmax(0,1fr))}
}
@media (max-width:720px){
  body{font-size:15px}
  .wrap{padding:0 16px}
  .marca h1{font-size:23px}
  .railmeta{text-align:left}
  .kpi .v{font-size:23px}
  .titular{padding:18px}
  .titular p{font-size:16.5px}
  .card{padding:15px 14px}
  .fetq{min-width:100%}
  .estado{margin-left:0;text-align:left}
  .accion .m,.accion .p{margin-left:0}
  .tabs .filtroact{display:none}
}
@media (prefers-reduced-motion:reduce){*{transition:none !important;animation:none !important}}
@media print{
  body{background:#fff}
  .tabs,.filtros,.railmeta .imp,.pager,.tablaTop .der,.leg .hint{display:none !important}
  .vista{display:block !important;break-before:page}
  .vista:first-of-type{break-before:auto}
  .card,.accion{break-inside:avoid;border-color:#CCC}
  details.lee{display:none}
  details.bloque{display:none}
  .rail{background:#fff !important;color:#000 !important}
  .marca h1,.railmeta b{color:#000 !important}
  .marca .planta,.marca .sub,.railmeta{color:#444 !important}
  .titular{background:#fff !important;border:1px solid #999}
  .titular p,.titular b{color:#000 !important}
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
  {k:'maquina', f:0, t:'Máquina', chips:true},
  {k:'modulo',  f:2, t:'Módulo'},
  {k:'familia', f:3, t:'Familia'},
  {k:'area',    f:5, t:'Área'},
  {k:'tipo',    f:4, t:'Tipo'},
  {k:'turno',   f:6, t:'Turno'},
  {k:'unidad',  f:1, t:'Unidad'}
];
var MESES = ['ene','feb','mar','abr','may','jun','jul','ago','set','oct','nov','dic'];
var MESL  = ['enero','febrero','marzo','abril','mayo','junio','julio','agosto',
             'setiembre','octubre','noviembre','diciembre'];
var QS   = ['CRITICO','CRONICO','AGUDO','LEVE'];
var QCOL = {CRITICO:'#B5352B', CRONICO:'#D07A1E', AGUDO:'#10607F', LEVE:'#93A7AF'};
var QNOM = {CRITICO:'Crítico', CRONICO:'Crónico', AGUDO:'Agudo', LEVE:'Leve'};
var QSUB = {CRITICO:'falla seguido y cada parada es larga',
            CRONICO:'falla seguido, pero se arregla rápido',
            AGUDO:'falla poco, pero cada parada es larga',
            LEVE:'falla poco y se arregla rápido'};
var QTXT = {
  CRITICO:'Pierde horas por las dos vías. Es por donde conviene empezar.',
  CRONICO:'El problema es que vuelve a fallar: se ataca con causa raíz, no con más rapidez.',
  AGUDO:'El problema es cuánto demora la reparación: repuesto, acceso y procedimiento.',
  LEVE:'Poco impacto en el periodo.'};
var TNOM = {SUBE:'Empeora', BAJA:'Mejora', IGUAL:'Sin cambio claro'};
var TCLS = {SUBE:'b-det', BAJA:'b-mej', IGUAL:'b-nd'};
var PAL = ['#10607F','#B5352B','#17714B','#D07A1E','#6B4B7A','#2E89AC','#9C6B00','#8A2E5D'];
var T975 = [null,12.706,4.303,3.182,2.776,2.571,2.447,2.365,2.306,2.262,2.228,
            2.201,2.179,2.160,2.145,2.131,2.120,2.110,2.101,2.093,2.086,
            2.080,2.074,2.069,2.064,2.060,2.056,2.052,2.048,2.045,2.042];
var DIA = 86400000;

/* ============================ estado ============================ */
var S = {
  sel:{maquina:[], unidad:[], modulo:[], familia:[], tipo:[], area:[], turno:[]},
  q:'', m0:null, m1:null, vista:'resumen', quadOff:{}, tab:{}, mas:false
};
var MESKEYS = [], COV = {};
// ventana de meses del filtro: VM = meses completos (para tendencias),
// REC = últimos meses completos, DIAS = días cubiertos, RANGO = fechas
var VM = [], REC = [], DIAS = 1, RANGO = {ini:0, fin:0};

/* ============================ utilidades ============================ */
function el(id){ return document.getElementById(id); }
function esc(s){ return String(s==null?'':s).replace(/[&<>"]/g, function(c){
  return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]; }); }
function fnum(v,d){ if(v==null||isNaN(v)) return '—';
  return Number(v).toLocaleString(M.idioma||'es-PE',
    {minimumFractionDigits:d, maximumFractionDigits:d}); }
function fh(v){ if(v==null||isNaN(v)) return '—';
  return fnum(v, v>=100?0:(v>=10?1:2)); }
function fmin(h){ if(h==null||isNaN(h)) return '—'; var m=h*60; return fnum(m, m<10?1:0); }
function fpc(v,d){ return fnum(v, d==null?1:d)+' %'; }
function tasa(v){ return fnum(v, v<10?1:0); }
function plural(n,a,b){ return n===1?a:b; }
function fechaObj(th){ return new Date(D.t0ms + th*3600000); }
function fechaCorta(th){ var d=fechaObj(th);
  return d.getUTCDate()+' '+MESES[d.getUTCMonth()]+' '+String(d.getUTCFullYear()).slice(2); }
function fechaHora(th){ var d=fechaObj(th), p=function(x){return (x<10?'0':'')+x;};
  return d.getUTCFullYear()+'-'+p(d.getUTCMonth()+1)+'-'+p(d.getUTCDate())+' '+
         p(d.getUTCHours())+':'+p(d.getUTCMinutes()); }
function fechaTxt(ms){ var d=new Date(ms);
  return d.getUTCDate()+' de '+MESL[d.getUTCMonth()]+' de '+d.getUTCFullYear(); }
function mesKeyMs(ms){ var d=new Date(ms); return d.getUTCFullYear()*12 + d.getUTCMonth(); }
function mesTxt(k){ return MESES[k%12]+' '+String(Math.floor(k/12)).slice(2); }
function mesLargo(k){ return MESL[k%12]+' de '+Math.floor(k/12); }
function mesMs(k){ return Date.UTC(Math.floor(k/12), k%12, 1); }
function modCod(i){ return D.dic.modulo[i]; }
function modLabel(i){ var n=D.modnom[i]; return n ? D.dic.modulo[i]+' — '+n : D.dic.modulo[i]; }
function dimTexto(k,i){ return k==='modulo' ? modLabel(i) : D.dic[k][i]; }
function actLabel(m,mo){ return D.dic.maquina[m]+' · '+D.dic.modulo[mo]; }
function actFull(m,mo){ return D.dic.maquina[m]+' · '+modLabel(mo); }
function suma(a){ var s=0; for(var i=0;i<a.length;i++) s+=a[i]; return s; }
function sumaH(rows){ var s=0; for(var i=0;i<rows.length;i++) s+=rows[i][F.H]; return s; }
function clamp(v,a,b){ return v<a?a:(v>b?b:v); }
function corta(s,n){ s=String(s); return s.length>n ? s.slice(0,n-1)+'…' : s; }

/* ============================ estadistica ============================ */
function t975(gl){ return gl < T975.length ? T975[gl] : 1.96 + 2.4/gl; }
/* ¿Fallan más (o menos) en la segunda mitad del periodo que en la primera?
   Igual que en el script Python. y: fallas por mes completo, e: días observados.
   Prueba de razón de tasas con la variación real de mes a mes dentro de cada
   mitad (cuasi-Poisson): un mes malo aislado no alcanza para decir que empeora. */
function tendencia(y, e){
  var k = y.length, i;
  if(k < M.minMeses || suma(y) < M.minTend) return null;
  var h = Math.floor(k/2), y1=0, y2=0, e1=0, e2=0;
  for(i=0;i<k;i++){ if(i<h){ y1+=y[i]; e1+=e[i]; } else { y2+=y[i]; e2+=e[i]; } }
  if(!(e1>0 && e2>0)) return null;
  var r1 = y1/e1, r2 = y2/e2, chi = 0, mu;
  for(i=0;i<k;i++){ mu = (i<h ? r1 : r2)*e[i]; if(mu>0) chi += (y[i]-mu)*(y[i]-mu)/mu; }
  var phi = Math.max(1, chi/(k-2)), n = y1+y2, p0 = e2/(e1+e2);
  var z = (y2 - n*p0)/Math.sqrt(phi*n*p0*(1-p0)), lim = t975(k-2);
  return {z:z, phi:phi, ini:r1*30.44, fin:r2*30.44,
          estado: z>lim ? 'SUBE' : (z<-lim ? 'BAJA' : 'IGUAL')};
}
function badgeT(t){
  if(!t) return '<span class="suave" title="Hacen falta al menos '+M.minTend+' fallas y '+
    M.minMeses+' meses completos en el filtro">—</span>';
  return '<span class="badge '+TCLS[t.estado]+'">'+TNOM[t.estado]+'</span>';
}
function badgeQ(q){ return '<span class="badge b-'+q.slice(0,3).toLowerCase()+'">'+QNOM[q]+'</span>'; }
function evidencia(t){
  return t ? tasa(t.ini)+' → '+tasa(t.fin) : '—';
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
  if(excepto!=='mes'){
    if(S.m0!=null && r[F.MK] < S.m0) return false;
    if(S.m1!=null && r[F.MK] > S.m1) return false;
  }
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
  REC = VM.length > M.recientes ? VM.slice(VM.length - M.recientes) : [];
}
function filtrosActivos(){
  var n = 0;
  DIMS.forEach(function(d){ n += S.sel[d.k].length; });
  if(S.q) n++;
  if(S.m0!==MESKEYS[0] || S.m1!==MESKEYS[MESKEYS.length-1]) n++;
  return n;
}

/* ============================ agrupacion ============================ */
function agrupa(rows, keyFn){
  var map = Object.create(null), ord = [], rec = {}, i, r, k, g;
  REC.forEach(function(m){ rec[m.k] = 1; });
  for(i=0;i<rows.length;i++){
    r = rows[i]; k = keyFn(r); g = map[k];
    if(!g){ g = map[k] = {key:k, n:0, h:0, r0:r, mods:Object.create(null), nmod:0,
                          mc:Object.create(null), hRec:0, nRec:0};
            ord.push(g); }
    g.n++; g.h += r[F.H];
    if(g.mods[r[F.MOD]]===undefined){ g.mods[r[F.MOD]]=0; g.nmod++; }
    g.mods[r[F.MOD]] += r[F.H];
    g.mc[r[F.MK]] = (g.mc[r[F.MK]]||0) + 1;
    if(rec[r[F.MK]]){ g.hRec += r[F.H]; g.nRec++; }
  }
  var e = VM.map(function(m){ return m.e; });
  for(i=0;i<ord.length;i++){
    g = ord[i];
    g.mttr = g.h/g.n;
    g.tend = tendencia(VM.map(function(m){ return g.mc[m.k]||0; }), e);
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
/* Grupos de Knights (2001): límite de frecuencia = fallas totales / equipos,
   límite de duración = horas totales / fallas totales (promedios, no medianas). */
function cuadrantes(acts){
  if(!acts.length) return {items:[], todos:[], ln:0, lm:0};
  var sn=0, sh=0;
  acts.forEach(function(g){ sn+=g.n; sh+=g.h; });
  var ln = sn/acts.length, lm = sh/sn;
  acts.forEach(function(g){
    g.q = (g.n>ln && g.mttr>lm) ? 'CRITICO' : (g.n>ln) ? 'CRONICO'
        : (g.mttr>lm) ? 'AGUDO' : 'LEVE'; });
  return {items: acts.filter(function(g){ return g.n >= M.minJK; }), todos:acts, ln:ln, lm:lm};
}
function filtraActivo(k){
  var p = String(k).split(/[|:]/);
  S.sel.maquina = [+p[0]]; S.sel.modulo = [+p[1]]; render();
}

/* ============================ tooltip ============================ */
var TIP;
function tipHtml(titulo, filas, pie){
  return '<div class="t">'+esc(titulo)+'</div>'+filas.map(function(f){
    return '<div class="r"><span>'+f[0]+'</span><b>'+f[1]+'</b></div>'; }).join('')+
    (pie ? '<div class="q">'+pie+'</div>' : '');
}
function tipShow(html, ev){
  if(!TIP) return;
  TIP.innerHTML = html; TIP.style.display='block'; tipMove(ev);
}
function tipMove(ev){
  if(!TIP || TIP.style.display==='none') return;
  var w = TIP.offsetWidth || 270, h = TIP.offsetHeight || 150;
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
                              p:0, per: cfg.per||12, q:''};
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
        return String(valor(f,c)).toLowerCase().indexOf(q) >= 0; }); });
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
  h.push('<div class="tablaTop"><div class="tt">'+esc(cfg.titulo||'')+
    (cfg.nota?' <span style="color:var(--gris2)">'+esc(cfg.nota)+'</span>':'')+
    '</div><div class="der">');
  if(cfg.buscar!==false && cfg.filas.length>12)
    h.push('<input type="search" class="tq" placeholder="Buscar en la tabla" value="'+
           esc(st.q)+'" style="width:168px">');
  if(cfg.csv!==false) h.push('<button class="btn g tcsv">Descargar CSV</button>');
  h.push('</div></div>');

  if(!filas.length){
    h.push('<div class="vacio"><b>'+esc(cfg.vacioT||'No hay datos con este filtro')+
      '</b>'+esc(cfg.vacio||'Quita alguna condición de la barra de filtros.')+'</div></div>');
    cont.innerHTML = h.join(''); enganchaTabla(cont, id, cfg); return;
  }
  h.push('<div class="scroll"><table><thead><tr>');
  cols.forEach(function(c,i){
    var cls = (c.tipo==='num'?'r ':'') + 's' +
      (st.c===i ? (st.asc?' asc':' desc') : '');
    h.push('<th class="'+cls+'" data-c="'+i+'" title="'+esc(c.ayuda||('Ordenar por '+c.t))+
      '">'+esc(c.t)+'<span class="ar"></span></th>');
  });
  h.push('</tr></thead><tbody>');
  vis.forEach(function(f){
    h.push('<tr'+(cfg.onRow?' class="click" data-k="'+esc(cfg.rowKey(f))+'"':'')+'>');
    cols.forEach(function(c,i){
      h.push('<td class="'+(c.tipo==='num'?'r ':'')+(c.cls||'')+'">'+
             celda(f,c,mx[i])+'</td>'); });
    h.push('</tr>');
  });
  h.push('</tbody></table></div>');
  if(filas.length > st.per){
    h.push('<div class="pager"><button class="btn g tprev"'+(st.p<=0?' disabled':'')+
      '>Anterior</button><button class="btn g tnext"'+(st.p>=tot-1?' disabled':'')+
      '>Siguiente</button><span>Filas '+fnum(st.p*st.per+1,0)+' a '+
      fnum(Math.min((st.p+1)*st.per, filas.length),0)+' de '+fnum(filas.length,0)+
      '</span><select class="tper" style="margin-left:auto">'+
      [10,25,50,100].map(function(n){ return '<option value="'+n+'"'+
        (n===st.per?' selected':'')+'>'+n+' por página</option>'; }).join('')+
      '</select></div>');
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
  var pv = box.querySelector('.tprev'), nx = box.querySelector('.tnext'),
      pp = box.querySelector('.tper');
  if(pv) pv.addEventListener('click', function(){ st.p--; tabla(cont,id,cfg); });
  if(nx) nx.addEventListener('click', function(){ st.p++; tabla(cont,id,cfg); });
  if(pp) pp.addEventListener('change', function(){ st.per=+pp.value; st.p=0; tabla(cont,id,cfg); });
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
  bajaCSV(csvTexto(cols.map(function(c){return c.t;}), filas),
          (cfg.archivo||'tabla')+'.csv');
}

/* ============================ helpers svg ============================ */
function ancho(cont, min){ return Math.max(min||280, cont.clientWidth || cont.offsetWidth || 880); }
function esMovil(){ return (window.innerWidth||1024) < 720; }
function vacio(cont, titulo, texto){
  cont.innerHTML = '<div class="vacio"><b>'+esc(titulo)+'</b>'+esc(texto||'')+'</div>';
}
function txtSvg(x,y,t,o){ o=o||{};
  return '<text x="'+x+'" y="'+y+'" font-size="'+(o.s||11)+'" fill="'+(o.c||'#5C7480')+
    '"'+(o.a?' text-anchor="'+o.a+'"':'')+(o.w?' font-weight="'+o.w+'"':'')+
    (o.op?' opacity="'+o.op+'"':'')+(o.tr?' transform="'+o.tr+'"':'')+
    ' pointer-events="none">'+esc(t)+'</text>'; }

/* ============================ barras horizontales ordenadas ============================ */
/* modo 'acum': muestra el % acumulado y marca el 80 % (Pareto).
   modo 'parte': muestra la parte de cada barra en el total. */
function gPareto(cont, grupos, opt){
  opt = opt || {};
  var g = porHoras(grupos), total = suma(g.map(function(x){return x.h;}));
  if(!g.length || total<=0){ vacio(cont,'Sin averías en esta selección','Ajusta los filtros.'); return; }
  var TOP = Math.min(opt.top || 14, g.length), movil = esMovil();
  var W = ancho(cont), etq = movil ? 104 : (opt.etq || 190), val = movil ? 70 : 118;
  var largo = Math.round(etq/7.2);
  var rh = 26, H = TOP*rh + 30;
  var iw = Math.max(40, W - etq - val - 14);
  var mx = g[0].h, acum = 0, corte = -1, modo = opt.modo || 'acum';
  var s = ['<svg viewBox="0 0 '+W+' '+H+'" width="100%" height="'+H+'">'], tips=[];
  for(var i=0;i<TOP;i++){
    var x = g[i], y = i*rh + 14, ancho_ = x.h/mx*iw;
    acum += x.h;
    var pa = acum/total*100;
    if(corte<0 && pa>=80) corte = i;
    tips.push(tipHtml(x.full, [
      ['Horas de avería', fh(x.h)+' h'],
      ['Averías', fnum(x.n,0)],
      ['Minutos por avería', fmin(x.mttr)],
      ['Parte del total', fpc(x.h/total*100)]].concat(modo==='acum' ?
      [['Acumulado hasta aquí', fpc(pa)]] : []), opt.click!==false?'Clic para filtrar':null));
    s.push('<g class="hit" data-t="'+(tips.length-1)+'"'+
      (opt.click!==false?' data-c="'+x.key+'"':'')+'>');
    s.push('<rect x="0" y="'+(y-11)+'" width="'+W+'" height="'+(rh-2)+'" fill="transparent"/>');
    s.push(txtSvg(etq-9, y+4, corta(x.label, largo), {s:12, c:'#1B3B4A', a:'end', w:500}));
    s.push('<rect x="'+etq+'" y="'+(y-8)+'" width="'+Math.max(1,ancho_).toFixed(1)+
      '" height="16" fill="'+(modo==='parte' || i<=corte || corte<0 ? '#10607F':'#93C6D8')+'" rx="1"/>');
    s.push(txtSvg(etq+ancho_+7, y+4, fh(x.h)+' h', {s:12, c:'#0B2430', w:600}));
    if(!movil) s.push(txtSvg(W-4, y+4, modo==='acum' ? fpc(pa,0)+' acum.' : fpc(x.h/total*100,0),
      {s:11, c:'#8199A3', a:'end'}));
    s.push('</g>');
  }
  if(modo==='acum' && corte>=0 && corte < TOP-1){
    var yc = (corte+1)*rh + 2;
    s.push('<line x1="0" y1="'+yc+'" x2="'+W+'" y2="'+yc+'" stroke="#B5352B" '+
      'stroke-width="1.2" stroke-dasharray="5 4"/>');
    s.push(txtSvg(etq, yc+13, 'hasta aquí se acumula el 80 % de las horas',
      {s:11, c:'#B5352B', w:600}));
  }
  s.push('</svg>');
  cont.innerHTML = s.join('');
  if(opt.click!==false) interactivo(cont, tips, opt.onClick || filtraActivo);
  else interactivo(cont, tips, null);
}

/* ============================ frecuencia contra duración (Jack-Knife) ============================ */
function gJK(cont, acts, res, leyenda){
  var items = res.items.filter(function(g){ return !S.quadOff[g.q]; });
  if(!items.length){
    vacio(cont, 'No hay equipos con al menos '+M.minJK+' fallas',
      'Amplía la selección o revisa los grupos ocultos.');
    if(leyenda) leyenda.innerHTML=''; return;
  }
  var movil = esMovil();
  var W = ancho(cont), H = movil ? 380 : Math.round(clamp(W*0.52, 360, 500));
  var Mg = {t:18, r:movil?18:56, b:46, l:56}, iw = W-Mg.l-Mg.r, ih = H-Mg.t-Mg.b;
  var xs = items.map(function(g){return g.n;}), ys = items.map(function(g){return g.mttr*60;});
  var x0 = Math.log10(Math.max(1, Math.min.apply(null,xs))*0.72),
      x1 = Math.log10(Math.max.apply(null,xs)*1.55),
      y0 = Math.log10(Math.max(0.05, Math.min.apply(null,ys))*0.62),
      y1 = Math.log10(Math.max.apply(null,ys)*2.0);
  if(x1-x0<0.35){ x0-=0.22; x1+=0.22; }
  if(y1-y0<0.35){ y0-=0.22; y1+=0.22; }
  var X = function(v){ return Mg.l + (Math.log10(v)-x0)/(x1-x0)*iw; };
  var Y = function(v){ return Mg.t + ih - (Math.log10(v)-y0)/(y1-y0)*ih; };
  var hmx = Math.max.apply(null, items.map(function(g){return g.h;}));
  var R = function(h){ return 4 + Math.sqrt(h/hmx)*(movil?14:20); };
  var s = ['<svg viewBox="0 0 '+W+' '+H+'" width="100%" height="'+H+'">'], tips=[];
  var ticks = function(a,b){ var o=[], e=Math.floor(a);
    for(; e<=Math.ceil(b); e++){ [1,2,5].forEach(function(m){
      var v=m*Math.pow(10,e); if(Math.log10(v)>=a && Math.log10(v)<=b) o.push(v); }); }
    return o; };
  ticks(x0,x1).forEach(function(v){
    s.push('<line x1="'+X(v)+'" y1="'+Mg.t+'" x2="'+X(v)+'" y2="'+(Mg.t+ih)+'" stroke="#E9EEEF"/>');
    s.push(txtSvg(X(v), Mg.t+ih+17, v, {a:'middle'})); });
  ticks(y0,y1).forEach(function(v){
    s.push('<line x1="'+Mg.l+'" y1="'+Y(v)+'" x2="'+(Mg.l+iw)+'" y2="'+Y(v)+'" stroke="#E9EEEF"/>');
    s.push(txtSvg(Mg.l-8, Y(v)+4, v<1?v.toFixed(1):v, {a:'end'})); });
  // diagonales: mismas horas totales
  [2,5,10,25,50,100,200,400].forEach(function(hh){
    var pts=[], n;
    for(n=Math.pow(10,x0); n<=Math.pow(10,x1); n*=1.07){
      var mt=hh*60/n, ly=Math.log10(mt);
      if(ly>=y0 && ly<=y1) pts.push(X(n).toFixed(1)+','+Y(mt).toFixed(1)); }
    if(pts.length>1){
      s.push('<polyline points="'+pts.join(' ')+'" fill="none" stroke="#D5DEE0" '+
        'stroke-width="1" stroke-dasharray="3 4"/>');
      var L = pts[pts.length-1].split(',');
      var cercaProm = Math.abs(+L[1] - Y(res.lm*60)) < 22 && +L[0] > Mg.l+iw-30;
      if(+L[1] > Mg.t+12 && !cercaProm)
        s.push(txtSvg(+L[0]+4, +L[1]-3, hh+' h', {s:9.5, c:'#A8B8BF'})); }
  });
  var lmMin = res.lm*60;
  if(res.ln>0 && lmMin>0){
    s.push('<line x1="'+X(res.ln)+'" y1="'+Mg.t+'" x2="'+X(res.ln)+'" y2="'+(Mg.t+ih)+
      '" stroke="#1B3B4A" stroke-width="1.2" opacity=".45"/>');
    s.push('<line x1="'+Mg.l+'" y1="'+Y(lmMin)+'" x2="'+(Mg.l+iw)+'" y2="'+Y(lmMin)+
      '" stroke="#1B3B4A" stroke-width="1.2" opacity=".45"/>');
    s.push(txtSvg(X(res.ln), Mg.t-5, 'promedio: '+fnum(res.ln,0)+' fallas', {s:10, c:'#5C7480', a:'middle'}));
    if(!movil){
      s.push(txtSvg(Mg.l+iw+5, Y(lmMin)-3, 'promedio', {s:10, c:'#5C7480'}));
      s.push(txtSvg(Mg.l+iw+5, Y(lmMin)+9, fnum(lmMin,0)+' min', {s:10, c:'#5C7480'}));
    }
  }
  [['CRITICO',Mg.l+iw-10,Mg.t+15,'end'],['CRONICO',Mg.l+iw-10,Mg.t+ih-8,'end'],
   ['AGUDO',Mg.l+10,Mg.t+15,'start'],['LEVE',Mg.l+10,Mg.t+ih-22,'start']].forEach(function(r){
    s.push(txtSvg(r[1], r[2], QNOM[r[0]].toUpperCase(), {s:13, w:700, c:QCOL[r[0]], op:.3, a:r[3]})); });
  s.push(txtSvg(Mg.l+iw/2, H-9, 'Cantidad de fallas en el periodo (escala log)', {s:12, c:'#1B3B4A', a:'middle'}));
  s.push(txtSvg(0, 0, 'Minutos por falla (escala log)', {s:12, c:'#1B3B4A', a:'middle',
    tr:'translate(15,'+(Mg.t+ih/2)+') rotate(-90)'}));

  items.sort(function(a,b){ return b.h-a.h; }).forEach(function(g){
    tips.push(tipHtml(g.full, [
      ['Fallas', fnum(g.n,0)],
      ['Horas totales', fh(g.h)+' h'],
      ['Minutos por falla', fmin(g.mttr)],
      ['Grupo', '<span style="color:'+QCOL[g.q]+'">'+QNOM[g.q]+'</span>'],
      ['Tendencia', g.tend ? TNOM[g.tend.estado] : 'pocos datos']
    ], QTXT[g.q]+'<br>Clic para filtrar'));
    s.push('<circle class="hit" data-t="'+(tips.length-1)+'" data-c="'+g.maq+':'+g.mod+
      '" cx="'+X(g.n).toFixed(1)+'" cy="'+Y(g.mttr*60).toFixed(1)+'" r="'+R(g.h).toFixed(1)+
      '" fill="'+QCOL[g.q]+'" fill-opacity=".55" stroke="#fff" stroke-width="1.5"/>');
  });
  // nombres de los equipos con más horas, sin encimarse
  var puestas = [];
  var choca = function(b){ return puestas.some(function(o){
    return b.x0 < o.x1 && b.x1 > o.x0 && b.y0 < o.y1 && b.y1 > o.y0; }); };
  items.slice(0, movil?5:9).forEach(function(g){
    var cx = X(g.n), cy = Y(g.mttr*60), r = R(g.h), w = g.label.length*6.3;
    var a = cx > Mg.l+iw-w/2 ? 'end' : (cx < Mg.l+w/2 ? 'start' : 'middle');
    var x0 = a==='end' ? cx-w : (a==='start' ? cx : cx-w/2);
    var opciones = [cy-r-5, cy+r+13];
    for(var o=0; o<opciones.length; o++){
      var ty = opciones[o];
      if(ty < Mg.t+11 || ty > Mg.t+ih-2) continue;
      var caja = {x0:x0-2, x1:x0+w+2, y0:ty-11, y1:ty+3};
      if(choca(caja)) continue;
      puestas.push(caja);
      s.push(txtSvg(cx, ty, g.label, {s:10.5, w:700, c:'#0B2430', a:a}));
      break;
    }
  });
  s.push('</svg>');
  cont.innerHTML = s.join('');
  interactivo(cont, tips, filtraActivo);

  if(leyenda){
    var cnt = {}; res.items.forEach(function(g){ cnt[g.q]=(cnt[g.q]||0)+1; });
    leyenda.innerHTML = QS.map(function(q){
      return '<button class="item'+(S.quadOff[q]?' off':'')+'" data-q="'+q+'" title="'+
        esc(QTXT[q])+'"><i style="background:'+QCOL[q]+'"></i>'+QNOM[q]+
        ' <b>'+(cnt[q]||0)+'</b></button>'; }).join('') +
      '<span class="hint">Clic en un grupo para ocultarlo. El tamaño del círculo son las horas.</span>';
    leyenda.querySelectorAll('button.item').forEach(function(b){
      b.addEventListener('click', function(){
        var q = b.getAttribute('data-q');
        S.quadOff[q] = !S.quadOff[q]; render(); }); });
  }
}

/* ============================ evolucion mensual ============================ */
function gMensual(cont, rows, leyenda){
  if(!rows.length){ vacio(cont,'Sin averías en esta selección',''); if(leyenda) leyenda.innerHTML=''; return; }
  var porMes = {}, tot = {}, i, r, k, m;
  for(i=0;i<rows.length;i++){
    r = rows[i]; k = r[F.MK]; m = r[F.MAQ];
    if(!porMes[k]) porMes[k] = {h:0, n:0, maq:{}};
    porMes[k].h += r[F.H]; porMes[k].n++;
    porMes[k].maq[m] = (porMes[k].maq[m]||0) + r[F.H];
    tot[m] = (tot[m]||0) + r[F.H];
  }
  var todos = MESKEYS.filter(function(kk){ return kk>=S.m0 && kk<=S.m1; });
  if(todos.length < 2){ vacio(cont,'Un solo mes en la selección',
    'Amplía el rango de meses para ver la evolución.'); if(leyenda) leyenda.innerHTML=''; return; }
  var top = Object.keys(tot).map(Number).sort(function(a,b){return tot[b]-tot[a];});
  var vis = top.slice(0, 6), resto = top.slice(6);
  var color = {}; vis.forEach(function(mm,ix){ color[mm] = PAL[ix%PAL.length]; });

  var W = ancho(cont), H = esMovil()? 280 : Math.round(clamp(W*0.36, 270, 360));
  var Mg = {t:16, r:12, b:40, l:50}, iw = W-Mg.l-Mg.r, ih = H-Mg.t-Mg.b;
  var mxV = 0;
  todos.forEach(function(kk){ if(porMes[kk] && porMes[kk].h > mxV) mxV = porMes[kk].h; });
  mxV = mxV*1.12 || 1;
  var paso = iw/todos.length, bw = Math.max(5, Math.min(54, paso - 8));
  var X = function(ix){ return Mg.l + (ix+0.5)*paso; };
  var Y = function(v){ return Mg.t + ih - v/mxV*ih; };
  var s = ['<svg viewBox="0 0 '+W+' '+H+'" width="100%" height="'+H+'">'], tips=[];
  var sug = M.sugerido ? (parseInt(M.sugerido.slice(0,4),10)*12 + parseInt(M.sugerido.slice(5,7),10)-1) : null;
  todos.forEach(function(kk, ix){
    var c = COV[kk], parcial = c.d < 0.5*c.dm, arranque = sug!=null && kk < sug;
    if(parcial || arranque){
      s.push('<rect x="'+(X(ix)-paso/2).toFixed(1)+'" y="'+Mg.t+'" width="'+paso.toFixed(1)+
        '" height="'+ih+'" fill="#8199A3" opacity=".10"/>');
      s.push(txtSvg(X(ix), Mg.t+12, parcial ? fnum(c.d,0)+' '+plural(Math.round(c.d),'día','días')
        : 'registro', {s:10, c:'#5C7480', a:'middle'}));
    }
  });
  for(var gg=0; gg<=4; gg++){
    var v = mxV*gg/4;
    s.push('<line x1="'+Mg.l+'" y1="'+Y(v)+'" x2="'+(Mg.l+iw)+'" y2="'+Y(v)+'" stroke="#E9EEEF"/>');
    s.push(txtSvg(Mg.l-7, Y(v)+4, fnum(v,0), {a:'end', s:10.5})); }
  s.push(txtSvg(0,0,'Horas de avería', {s:11.5, c:'#1B3B4A', a:'middle',
    tr:'translate(13,'+(Mg.t+ih/2)+') rotate(-90)'}));
  todos.forEach(function(kk, ix){
    var mm = porMes[kk];
    if(ix % (todos.length>15 ? 2 : 1) === 0)
      s.push(txtSvg(X(ix), Mg.t+ih+17, mesTxt(kk), {a:'middle', s:10.5}));
    if(!mm) return;
    var y = Y(0), det = [];
    vis.forEach(function(mq){
      var v2 = mm.maq[mq]||0; if(v2<=0) return;
      var hh = v2/mxV*ih; y -= hh; det.push([mq, v2]);
      s.push('<rect x="'+(X(ix)-bw/2).toFixed(1)+'" y="'+y.toFixed(1)+'" width="'+bw.toFixed(1)+
        '" height="'+Math.max(0.6,hh).toFixed(1)+'" fill="'+color[mq]+'"/>'); });
    var vr = 0; resto.forEach(function(mq){ vr += mm.maq[mq]||0; });
    if(vr>0){ var hr = vr/mxV*ih; y -= hr;
      s.push('<rect x="'+(X(ix)-bw/2).toFixed(1)+'" y="'+y.toFixed(1)+'" width="'+bw.toFixed(1)+
        '" height="'+Math.max(0.6,hr).toFixed(1)+'" fill="#C3D2D7"/>'); }
    s.push(txtSvg(X(ix), y-5, fnum(mm.h,0), {a:'middle', s:10.5, c:'#34586A', w:600}));
    var filas = det.sort(function(a,b){return b[1]-a[1];}).map(function(o){
      return ['<i class="dot" style="background:'+color[o[0]]+'"></i>'+esc(D.dic.maquina[o[0]]),
              fh(o[1])+' h']; });
    if(vr>0) filas.push(['<i class="dot" style="background:#C3D2D7"></i>Otras máquinas', fh(vr)+' h']);
    var c2 = COV[kk];
    tips.push(tipHtml(mesLargo(kk),
      [['Horas del mes', fh(mm.h)+' h'], ['Averías', fnum(mm.n,0)],
       ['Horas por día', fh(mm.h/Math.max(c2.d,1/24))]].concat(filas),
      c2.d < 0.5*c2.dm ? 'Mes incompleto: '+fnum(c2.d,1)+' días con datos.' : 'Clic para ver solo este mes'));
    s.push('<rect class="hit" data-t="'+(tips.length-1)+'" data-c="'+kk+'" x="'+
      (X(ix)-paso/2).toFixed(1)+'" y="'+Mg.t+'" width="'+paso.toFixed(1)+
      '" height="'+ih+'" fill="transparent"/>');
  });
  s.push('</svg>');
  cont.innerHTML = s.join('');
  interactivo(cont, tips, function(k2){ S.m0 = S.m1 = +k2; render(); });
  if(leyenda) leyenda.innerHTML = vis.map(function(mq){
    return '<span class="item"><i style="background:'+color[mq]+'"></i>'+
      esc(D.dic.maquina[mq])+'</span>'; }).join('') +
    (resto.length ? '<span class="item"><i style="background:#C3D2D7"></i>Otras '+
      resto.length+' máquinas</span>' : '') +
    '<span class="hint">Clic en un mes para ver solo ese mes. En gris, meses con pocos días de datos.</span>';
}

/* ============================ matriz maquina x modulo ============================ */
function gMatriz(cont, rows){
  if(!rows.length){ vacio(cont,'Sin averías en esta selección',''); return; }
  var maqs = porHoras(porMaquina(rows)).slice(0,12);
  var hMod = {};
  rows.forEach(function(r){ hMod[r[F.MOD]] = (hMod[r[F.MOD]]||0) + r[F.H]; });
  var mods = Object.keys(hMod).map(Number).sort(function(a,b){return hMod[b]-hMod[a];})
               .slice(0, esMovil()?8:13);
  if(!mods.length){ vacio(cont,'Sin módulos en esta selección',''); return; }
  var celdas = {};
  rows.forEach(function(r){
    var k = r[F.MAQ]+':'+r[F.MOD];
    if(!celdas[k]) celdas[k] = {h:0, n:0};
    celdas[k].h += r[F.H]; celdas[k].n++; });
  var cw = 62, rh = 30, lw = 66, th = 56, tw = 62;
  var W = Math.max(ancho(cont), lw + mods.length*cw + tw);
  var H = th + maqs.length*rh + 16;
  var vmax = 0;
  maqs.forEach(function(g){ mods.forEach(function(mo){
    var c = celdas[g.maq+':'+mo]; if(c && c.h>vmax) vmax = c.h; }); });
  var s = ['<svg viewBox="0 0 '+W+' '+H+'" width="'+W+'" height="'+H+'">'], tips=[];
  mods.forEach(function(mo, j){
    var x = lw + j*cw;
    s.push(txtSvg(x+cw/2, th-32, modCod(mo), {s:11.5, w:600, c:'#0B2430', a:'middle'}));
    var nm = D.modnom[mo] || '';
    if(nm) s.push(txtSvg(x+cw/2, th-19, corta(nm, 11), {s:9.5, c:'#8199A3', a:'middle'}));
  });
  s.push(txtSvg(lw + mods.length*cw + tw/2, th-32, 'Total', {s:11.5, w:600, c:'#0B2430', a:'middle'}));
  maqs.forEach(function(g, i){
    var y = th + i*rh;
    s.push(txtSvg(lw-9, y+rh/2+4, D.dic.maquina[g.maq], {s:12, w:600, c:'#0B2430', a:'end'}));
    mods.forEach(function(mo, j){
      var c = celdas[g.maq+':'+mo], x = lw + j*cw;
      var v = c ? c.h : 0;
      var t = vmax>0 ? Math.pow(v/vmax, 0.62) : 0;
      var fill = v>0 ? 'rgb('+Math.round(255-165*t)+','+Math.round(255-200*t)+','+
                 Math.round(255-210*t)+')' : '#FBFCFC';
      tips.push(v>0 ? tipHtml(D.dic.maquina[g.maq]+' · '+modLabel(mo), [
          ['Horas de avería', fh(v)+' h'],
          ['Averías', fnum(c.n,0)],
          ['Minutos por avería', fmin(v/c.n)],
          ['Parte de la máquina', fpc(v/g.h*100)]], 'Clic para filtrar')
        : tipHtml(D.dic.maquina[g.maq]+' · '+modLabel(mo), [['Horas de avería','0 h']],
          'Este módulo no registró averías en esta máquina'));
      s.push('<rect class="hit" data-t="'+(tips.length-1)+'"'+(v>0?' data-c="'+g.maq+':'+mo+'"':'')+
        ' x="'+(x+1)+'" y="'+(y+1)+'" width="'+(cw-2)+'" height="'+(rh-2)+'" fill="'+fill+
        '" stroke="#E9EEEF"/>');
      if(v>0) s.push(txtSvg(x+cw/2, y+rh/2+4, v>=10?fnum(v,0):fnum(v,1),
        {s:11, c: t>0.62?'#FFFFFF':'#1B3B4A', a:'middle', w: t>0.42?600:400}));
    });
    var xt = lw + mods.length*cw;
    s.push('<rect x="'+(xt+1)+'" y="'+(y+1)+'" width="'+(tw-2)+'" height="'+(rh-2)+
      '" fill="#EFF4F5" stroke="#E9EEEF"/>');
    s.push(txtSvg(xt+tw/2, y+rh/2+4, fh(g.h), {s:11.5, w:600, c:'#0B2430', a:'middle'}));
  });
  s.push('</svg>');
  cont.innerHTML = '<div class="scroll">'+s.join('')+'</div>';
  interactivo(cont, tips, filtraActivo);
}

/* ============================ barras apiladas por máquina ============================ */
function gApilado(cont, rows, dimKey, leyenda){
  var dim = DIMS.filter(function(d){return d.k===dimKey;})[0];
  if(!rows.length){ vacio(cont,'Sin averías en esta selección',''); if(leyenda) leyenda.innerHTML=''; return; }
  var maqs = porHoras(porMaquina(rows)).slice(0,10);
  var cat = {};
  rows.forEach(function(r){ cat[r[dim.f]] = (cat[r[dim.f]]||0) + r[F.H]; });
  var cats = Object.keys(cat).map(Number).sort(function(a,b){return cat[b]-cat[a];});
  var vis = cats.slice(0,7), resto = cats.slice(7);
  var color = {}; vis.forEach(function(c,i){ color[c] = PAL[i%PAL.length]; });
  var datos = {};
  rows.forEach(function(r){
    var k = r[F.MAQ]+':'+(vis.indexOf(r[dim.f])>=0 ? r[dim.f] : 'x');
    datos[k] = (datos[k]||0) + r[F.H]; });
  var tot = suma(maqs.map(function(g){ return g.h; })) || 1;
  var W = ancho(cont), etq = esMovil()?52:64, val = esMovil()?84:104;
  var rh = 28, H = maqs.length*rh + 8, iw = Math.max(40, W-etq-val);
  var mx = maqs.length ? maqs[0].h : 1;
  var s = ['<svg viewBox="0 0 '+W+' '+H+'" width="100%" height="'+H+'">'], tips=[];
  maqs.forEach(function(g,i){
    var y = i*rh + 5, x = etq;
    s.push(txtSvg(etq-8, y+14, D.dic.maquina[g.maq], {s:12, w:600, c:'#0B2430', a:'end'}));
    var lista = vis.map(function(c){ return [c, datos[g.maq+':'+c]||0]; })
                   .filter(function(o){ return o[1]>0; });
    var vr = datos[g.maq+':x']||0;
    lista.forEach(function(o){
      var w = o[1]/mx*iw;
      tips.push(tipHtml(D.dic.maquina[g.maq]+' · '+dimTexto(dimKey,o[0]), [
        ['Horas de avería', fh(o[1])+' h'],
        ['Parte de la máquina', fpc(o[1]/g.h*100)]], 'Clic para filtrar'));
      s.push('<rect class="hit" data-t="'+(tips.length-1)+'" data-c="'+g.maq+':'+o[0]+
        '" x="'+x.toFixed(1)+'" y="'+y+'" width="'+Math.max(0.8,w).toFixed(1)+
        '" height="18" fill="'+color[o[0]]+'"/>');
      x += w; });
    if(vr>0){ var w2 = vr/mx*iw;
      tips.push(tipHtml(D.dic.maquina[g.maq]+' · otras categorías',
        [['Horas de avería', fh(vr)+' h']]));
      s.push('<rect class="hit" data-t="'+(tips.length-1)+'" x="'+x.toFixed(1)+'" y="'+y+
        '" width="'+Math.max(0.8,w2).toFixed(1)+'" height="18" fill="#C3D2D7"/>'); x += w2; }
    s.push(txtSvg(x+7, y+14, fh(g.h)+' h · '+fpc(g.h/tot*100,0), {s:11.5, w:600, c:'#0B2430'}));
  });
  s.push('</svg>');
  cont.innerHTML = s.join('');
  interactivo(cont, tips, function(k){
    var p = k.split(':'); S.sel.maquina=[+p[0]]; S.sel[dimKey]=[+p[1]]; render(); });
  if(leyenda) leyenda.innerHTML = vis.map(function(c){
    return '<span class="item"><i style="background:'+color[c]+'"></i>'+
      esc(dimTexto(dimKey,c))+' <b>'+fpc(cat[c]/(suma(cats.map(function(x){return cat[x];}))||1)*100,0)+
      '</b></span>'; }).join('') +
    (resto.length?'<span class="item"><i style="background:#C3D2D7"></i>Otras '+
      resto.length+'</span>':'');
}

/* ============================ sin codificar ============================ */
var RE_SC = /LIBRE|SIN CLASIF|NO CLASIF|SIN DATO|SIN MAPEAR/i;
var SC = {};
function marcaSinCod(){
  ['l3','l4','familia','modulo'].forEach(function(k){
    SC[k] = D.dic[k].map(function(v){ return RE_SC.test(v); }); });
}

/* ============================ barra de filtros ============================ */
function cuenta(dimKey){
  var d = DIMS.filter(function(x){return x.k===dimKey;})[0];
  var base = filtra(dimKey), c = {}, i;
  for(i=0;i<base.length;i++) c[base[i][d.f]] = (c[base[i][d.f]]||0) + 1;
  return c;
}
function pintaChips(){
  var c = cuenta('maquina'), sel = S.sel.maquina;
  var orden = D.dic.maquina.map(function(v,i){ return i; })
    .sort(function(a,b){ return (c[b]||0)-(c[a]||0) || D.dic.maquina[a].localeCompare(D.dic.maquina[b]); });
  var h = orden.map(function(i){
    var on = sel.indexOf(i)>=0, hay = (c[i]||0) > 0;
    return '<button class="chip'+(hay?'':' na')+'" aria-pressed="'+on+'" data-i="'+i+'"'+
      (hay?'':' title="Sin averías con los demás filtros"')+'>'+esc(D.dic.maquina[i])+
      '<span class="n">'+fnum(c[i]||0,0)+'</span></button>'; }).join('');
  if(sel.length) h += '<button class="chip" data-i="all" style="border-style:dashed">Todas</button>';
  var cont = el('fchips'); cont.innerHTML = h;
  cont.querySelectorAll('.chip').forEach(function(b){
    b.addEventListener('click', function(){
      var v = b.getAttribute('data-i');
      if(v==='all'){ S.sel.maquina = []; }
      else { var i = +v, p = S.sel.maquina.indexOf(i);
             if(p>=0) S.sel.maquina.splice(p,1); else S.sel.maquina.push(i); }
      render(); }); });
}
var ABIERTO = null, MSQ = {};
function pintaSelects(){
  var cont = el('fsel');
  if(!cont.children.length){
    cont.innerHTML = DIMS.filter(function(d){return !d.chips;}).map(function(d){
      return '<div class="msel" data-k="'+d.k+'"><button type="button"></button>'+
             '<div class="pop"></div></div>'; }).join('');
    cont.querySelectorAll('.msel').forEach(function(m){
      m.querySelector('button').addEventListener('click', function(e){
        e.stopPropagation();
        var k = m.getAttribute('data-k');
        ABIERTO = (ABIERTO===k) ? null : k;
        pintaSelects(); }); });
  }
  cont.querySelectorAll('.msel').forEach(function(m){
    var k = m.getAttribute('data-k');
    var d = DIMS.filter(function(x){return x.k===k;})[0];
    var sel = S.sel[k], b = m.querySelector('button');
    b.textContent = sel.length===0 ? d.t+': todos'
      : sel.length===1 ? d.t+': '+corta(dimTexto(k, sel[0]), 20)
      : d.t+': '+sel.length+' elegidos';
    b.className = sel.length ? 'act' : '';
    m.className = 'msel' + (ABIERTO===k ? ' abierto' : '');
    if(ABIERTO===k) pintaPop(m, k); else m.querySelector('.pop').innerHTML='';
  });
}
function pintaPop(m, k){
  var c = cuenta(k), pop = m.querySelector('.pop'), sel = S.sel[k];
  var q = (MSQ[k]||'').toLowerCase();
  var idx = D.dic[k].map(function(v,i){return i;}).filter(function(i){
    return !q || dimTexto(k,i).toLowerCase().indexOf(q)>=0; });
  idx.sort(function(a,b){ return (c[b]||0)-(c[a]||0) ||
    dimTexto(k,a).localeCompare(dimTexto(k,b),'es'); });
  var h = [];
  if(D.dic[k].length > 8)
    h.push('<input type="search" placeholder="Buscar" value="'+esc(MSQ[k]||'')+'">');
  h.push('<div class="lista">');
  idx.slice(0,400).forEach(function(i){
    var on = sel.indexOf(i)>=0, hay = (c[i]||0)>0;
    h.push('<label class="'+(hay?'':'na')+'"><input type="checkbox" data-i="'+i+'"'+
      (on?' checked':'')+'><span>'+esc(dimTexto(k,i))+'</span><span class="c">'+
      fnum(c[i]||0,0)+'</span></label>'); });
  if(!idx.length) h.push('<div style="padding:10px;color:var(--gris);font-size:13.4px">'+
    'Nada coincide con la búsqueda.</div>');
  h.push('</div><div class="pie"><button data-a="none">Quitar selección</button>'+
    '<button data-a="close">Listo</button></div>');
  pop.innerHTML = h.join('');
  var inp = pop.querySelector('input[type=search]');
  if(inp){
    inp.addEventListener('input', function(){ MSQ[k]=inp.value; pintaPop(m,k);
      var n = m.querySelector('input[type=search]');
      if(n){ n.focus(); try{ n.setSelectionRange(n.value.length,n.value.length);}catch(e){} } });
    inp.addEventListener('click', function(e){ e.stopPropagation(); });
  }
  pop.querySelectorAll('input[type=checkbox]').forEach(function(ch){
    ch.addEventListener('change', function(){
      var i = +ch.getAttribute('data-i'), p = S.sel[k].indexOf(i);
      if(ch.checked && p<0) S.sel[k].push(i);
      if(!ch.checked && p>=0) S.sel[k].splice(p,1);
      render(); }); });
  pop.querySelectorAll('.pie button').forEach(function(b){
    b.addEventListener('click', function(e){
      e.stopPropagation();
      if(b.getAttribute('data-a')==='none'){ S.sel[k]=[]; render(); }
      else { ABIERTO=null; pintaSelects(); } }); });
  pop.addEventListener('click', function(e){ e.stopPropagation(); });
}
function pintaMeses(){
  var s0 = el('fm0'), s1 = el('fm1');
  if(!s0.options.length){
    var op = MESKEYS.map(function(k){ var c = COV[k];
      return '<option value="'+k+'">'+mesLargo(k)+
        (c.d < 0.5*c.dm ? ' ('+fnum(c.d,0)+' '+plural(Math.round(c.d),'día','días')+')' : '')+
        '</option>'; }).join('');
    s0.innerHTML = op; s1.innerHTML = op;
  }
  s0.value = String(S.m0); s1.value = String(S.m1);
}
function pintaPills(){
  var p = [];
  DIMS.forEach(function(d){
    S.sel[d.k].forEach(function(i){
      p.push(['<b>'+d.t+'</b> '+esc(corta(dimTexto(d.k,i),26)), d.k+':'+i]); }); });
  if(S.m0!==MESKEYS[0] || S.m1!==MESKEYS[MESKEYS.length-1])
    p.push(['<b>Meses</b> '+mesTxt(S.m0)+(S.m1!==S.m0?' a '+mesTxt(S.m1):''), 'mes:0']);
  if(S.q) p.push(['<b>Texto</b> '+esc(S.q), 'q:0']);
  var cont = el('fpills');
  cont.classList.toggle('oculto', !p.length);
  cont.innerHTML = p.map(function(o){
    return '<button class="chip" aria-pressed="true" data-k="'+o[1]+'" '+
      'title="Quitar este filtro">'+o[0]+' <span style="opacity:.7">✕</span></button>'; }).join('') +
    (p.length>1 ? '<button class="btn g" data-k="todo">Quitar todos</button>' : '');
  cont.querySelectorAll('[data-k]').forEach(function(b){
    b.addEventListener('click', function(){
      var p2 = b.getAttribute('data-k').split(':'), k = p2[0];
      if(k==='todo'){ limpiar(); return; }
      if(k==='mes'){ S.m0=MESKEYS[0]; S.m1=MESKEYS[MESKEYS.length-1]; }
      else if(k==='q'){ S.q=''; el('fq').value=''; }
      else { var i = S.sel[k].indexOf(+p2[1]); if(i>=0) S.sel[k].splice(i,1); }
      render(); }); });
  var n = filtrosActivos(), t = el('tfiltro');
  if(t) t.textContent = n ? 'Con '+n+' '+plural(n,'filtro','filtros')+' activos' : '';
}
function pintaEstado(rows){
  var h = sumaH(rows);
  el('festado').innerHTML = '<b>'+fnum(rows.length,0)+'</b> '+
    plural(rows.length,'avería','averías')+' · <b>'+fh(h)+' h</b>' +
    (rows.length < D.rows.length ? ' <span style="color:var(--gris2)">de '+
      fnum(D.rows.length,0)+'</span>' : '');
}
function pintaMas(){
  var n = 0;
  DIMS.forEach(function(d){ if(!d.chips) n += S.sel[d.k].length; });
  if(S.q) n++;
  var b = el('fmas');
  b.textContent = (S.mas ? 'Menos filtros' : 'Más filtros') + (n ? ' ('+n+')' : '');
  b.setAttribute('aria-expanded', String(S.mas));
  el('fmasbox').classList.toggle('oculto', !S.mas);
}
function limpiar(){
  DIMS.forEach(function(d){ S.sel[d.k] = []; });
  S.q=''; el('fq').value=''; S.m0=MESKEYS[0]; S.m1=MESKEYS[MESKEYS.length-1];
  S.quadOff={}; MSQ={}; ABIERTO=null; render();
}
function presets(){
  var ult = MESKEYS[MESKEYS.length-1], comp = MESKEYS.filter(function(k){
    return COV[k].d >= 0.5*COV[k].dm; }), p = [['Todo el periodo', MESKEYS[0], ult]];
  if(M.sugerido){
    var sk = parseInt(M.sugerido.slice(0,4),10)*12 + parseInt(M.sugerido.slice(5,7),10)-1;
    if(sk > MESKEYS[0] && sk <= ult) p.push(['Registro estable', sk, ult]);
  }
  [3,6].forEach(function(n){
    if(comp.length > n) p.push(['Últimos '+n+' meses', comp[comp.length-n], ult]); });
  el('fpresets').innerHTML = p.map(function(o,i){
    return '<button class="chip" data-i="'+i+'">'+esc(o[0])+'</button>'; }).join('');
  el('fpresets').querySelectorAll('.chip').forEach(function(b){
    b.addEventListener('click', function(){
      var o = p[+b.getAttribute('data-i')];
      S.m0 = o[1]; S.m1 = o[2]; render(); }); });
}

/* ============================ RESUMEN ============================ */
function mesesRecTxt(){
  if(!REC.length) return '';
  return REC.length===1 ? mesLargo(REC[0].k)
    : 'los últimos '+REC.length+' meses ('+mesTxt(REC[0].k)+' a '+mesTxt(REC[REC.length-1].k)+')';
}
function pintaTitular(rows, acts, maqs){
  var c = el('r-titular');
  if(!rows.length){ c.innerHTML = '<p>No hay averías con los filtros actuales. '+
    'Quita alguna condición para volver a ver datos.</p>'; return; }
  var tot = sumaH(rows), m0 = porHoras(maqs)[0], tipos = porHoras(porDim(rows,'tipo'));
  var t = 'Entre el '+fechaTxt(RANGO.ini)+' y el '+fechaTxt(RANGO.fin-1)+' se perdieron <b>'+
    fh(tot)+' horas</b> por averías'+(maqs.length>1 ? ' en '+maqs.length+' máquinas' :
    ' en '+esc(m0.label))+': <b>'+fh(tot/DIAS)+' h por día</b>.';
  if(maqs.length>1) t += ' La que más pierde es <b>'+esc(m0.label)+'</b>, con el '+
    fpc(m0.h/tot*100,0)+'.';
  if(tipos.length>1 && tipos[0].h/tot >= 0.55) t += ' Las '+esc(tipos[0].label.toLowerCase())+
    ' explican el '+fpc(tipos[0].h/tot*100,0)+' de las horas.';
  var p2 = '';
  if(acts.length >= 15){
    var ord = porHoras(acts), p10 = suma(ord.slice(0,10).map(function(g){return g.h;}))/tot*100;
    p2 = p10 < 50
      ? 'Las horas están repartidas: los 10 equipos que más pierden suman solo el '+fpc(p10,0)+
        '. Además de atacarlos uno por uno, conviene buscar las fallas que se repiten en muchas '+
        'máquinas (gráfico «Qué falla»).'
      : 'Las horas están concentradas: 10 equipos suman el '+fpc(p10,0)+
        '. Atacarlos primero es lo que más rinde.';
  }
  c.innerHTML = '<p>'+t+'</p>'+(p2 ? '<p>'+p2+'</p>' : '');
}
function pintaKPIs(cont, rows){
  var h = sumaH(rows), n = rows.length;
  var k = function(v,u,t,ay){ return '<div class="kpi"'+(ay?' title="'+esc(ay)+'"':'')+
    '><span class="v">'+v+(u?'<i>'+u+'</i>':'')+'</span><div class="k">'+t+'</div></div>'; };
  cont.innerHTML =
    k(fh(h),'h','Horas de avería') +
    k(fnum(n,0),'','Averías', 'Los registros que eran partes de una misma parada cuentan como una sola avería') +
    k(n?fmin(h/n):'—','min','Duración promedio de cada avería') +
    k(fh(h/DIAS),'h','Horas de avería por día', 'Horas de avería divididas por los días del periodo, sumando todas las máquinas');
}
function paso(g){
  var t = g.tend, s;
  if(t && t.estado==='BAJA')
    return 'Ya viene mejorando (de '+tasa(t.ini)+' a '+tasa(t.fin)+
      ' fallas por mes): confirmar qué se hizo y sostenerlo.';
  if(t && t.estado==='SUBE')
    s = 'Las fallas van en aumento (de '+tasa(t.ini)+' a '+tasa(t.fin)+' por mes): revisar qué '+
      'cambió (repuestos, ajustes, operación) y atacar la causa del síntoma principal.';
  else if(g.q==='CRITICO') s = 'Falla seguido y cada parada es larga: analizar la causa del '+
    'síntoma principal y por qué la reparación demora.';
  else if(g.q==='CRONICO') s = 'Falla seguido pero se arregla rápido: el foco es que no vuelva '+
    'a fallar (causa raíz del síntoma principal).';
  else if(g.q==='AGUDO') s = 'Falla poco, pero cada parada es larga: el foco es reparar más '+
    'rápido (repuesto a mano, acceso, procedimiento).';
  else s = 'Impacto moderado: seguimiento.';
  return s;
}
function pintaAcciones(cont, rows, acts){
  if(!acts.length){ vacio(cont, 'Sin averías en esta selección', ''); return; }
  var rec = REC.length > 0;
  el('r-accsub').innerHTML = rec
    ? 'Los '+Math.min(M.topAcc, acts.length)+' equipos que más horas perdieron en '+
      mesesRecTxt()+'. Un equipo es un módulo de una máquina, por ejemplo PF4 · CS. '+
      'Clic en uno para filtrar todo el reporte por él.'
    : 'Los equipos que más horas perdieron en el periodo elegido. Un equipo es un módulo de '+
      'una máquina, por ejemplo PF4 · CS. Clic en uno para filtrar todo el reporte por él.';
  // un equipo sin módulo no se puede asignar a nadie: no entra en la lista
  var lista = acts.filter(function(g){ return D.dic.modulo[g.mod] !== '(sin dato)'; })
    .sort(function(a,b){ return rec ? (b.hRec-a.hRec) || (b.h-a.h) : b.h-a.h; })
    .slice(0, M.topAcc);
  if(!lista.length){ vacio(cont, 'Ninguna avería de esta selección tiene módulo',
    'Sin módulo no se puede decir qué equipo atender.'); return; }
  var sint = {};
  lista.forEach(function(g){ sint[g.key] = {}; });
  rows.forEach(function(r){ var s2 = sint[r[F.MAQ]+'|'+r[F.MOD]];
    if(s2) s2[r[F.L4]] = (s2[r[F.L4]]||0) + 1; });
  cont.innerHTML = lista.map(function(g, i){
    var cnt = sint[g.key], sx = Object.keys(cnt).map(Number)
      .sort(function(a,b){ return cnt[b]-cnt[a]; });
    var s0 = sx[0], gen = !!D.gen[s0];
    var rep = '«'+esc(D.dic.l4[s0])+'» en '+fnum(cnt[s0],0)+' de '+fnum(g.n,0)+' fallas'+
      (gen ? ' <span class="gen">(detalle genérico)</span>' : '');
    var m = (rec ? '<b>'+fh(g.hRec)+' h</b> en esos meses · ' : '') +
      fh(g.h)+' h en todo el periodo · '+fnum(g.n,0)+' fallas · '+fmin(g.mttr)+' min por falla';
    return '<div class="accion '+g.q+'" data-k="'+g.maq+':'+g.mod+'" role="button" tabindex="0">'+
      '<div class="cab"><span class="pos">'+(i+1)+'</span><span class="t">'+esc(g.full)+'</span>'+
      '<span class="tags">'+badgeQ(g.q)+
      (g.tend && g.tend.estado!=='IGUAL' ? ' '+badgeT(g.tend) : '')+'</span></div>'+
      '<div class="m">'+m+'</div>'+
      '<div class="p"><b>Lo que más se repite:</b> '+rep+'</div>'+
      '<div class="p"><b>Siguiente paso:</b> '+paso(g)+'</div></div>';
  }).join('') + (function(){
    var ng = lista.filter(function(g){ var c=sint[g.key], mx=-1, s0=null;
      Object.keys(c).forEach(function(x){ if(c[x]>mx){ mx=c[x]; s0=+x; } });
      return D.gen[s0]; }).length;
    return ng ? '<p class="suave" style="margin:10px 0 0;font-size:13.6px">En '+ng+' de estos '+
      lista.length+' equipos el detalle que más se repite es genérico: antes de buscar la causa '+
      'hay que precisar qué pieza falla (órdenes de trabajo o el técnico que reparó).</p>' : '';
  })();
  cont.querySelectorAll('.accion').forEach(function(a){
    var ir = function(){ filtraActivo(a.getAttribute('data-k')); };
    a.addEventListener('click', ir);
    a.addEventListener('keydown', function(e){ if(e.key==='Enter') ir(); });
  });
}
function pintaAvisos(cont){
  var a = [], ult = MESKEYS[MESKEYS.length-1], cu = COV[ult];
  if(M.pctGenericoH >= 20) a.push('En el <b>'+fpc(M.pctGenericoH,0)+' de las horas</b> el '+
    'detalle (L4) es genérico, como «Regulación electrónica» o «Falla general»: dice que hubo '+
    'que intervenir, pero no qué se rompió.');
  if(M.pctSinCodH >= 5) a.push('El <b>'+fpc(M.pctSinCodH,0)+' de las horas</b> figura como '+
    '«Libre» o sin clasificar'+(M.peorCod && M.peorCod[1] >= 30 ? ', y se concentra en <b>'+
    esc(M.peorCod[0])+'</b>: ahí es el '+fpc(M.peorCod[1],0)+' de sus horas, así que sus '+
    'fallas no se pueden diagnosticar bien.' : '.'));
  if(cu && cu.d < 0.5*cu.dm) a.push(mesLargo(ult).replace(/^./, function(x){
    return x.toUpperCase(); })+' tiene solo '+fnum(cu.d,0)+' '+plural(Math.round(cu.d),'día','días')+
    ' de datos: no lo compares con los demás meses.');
  cont.innerHTML = a.length ? '<div class="nota"><b>Antes de decidir con estos números</b><ul>'+
    a.map(function(x){ return '<li>'+x+'</li>'; }).join('')+
    '</ul><p style="margin-top:6px">Más detalle en la pestaña Notas.</p></div>' : '';
}
function vResumen(rows){
  var acts = porActivo(rows), maqs = porMaquina(rows);
  cuadrantes(acts);
  pintaTitular(rows, acts, maqs);
  pintaKPIs(el('r-kpis'), rows);
  gApilado(el('r-maq'), rows, 'tipo', el('r-maqleg'));
  gPareto(el('r-fam'), porDim(rows,'familia'), {top:8, modo:'parte', etq:230,
    onClick:function(k){ S.sel.familia=[+k]; render(); }});
  pintaAcciones(el('r-acciones'), rows, acts);
  pintaAvisos(el('r-avisos'));
}

/* ============================ MAQUINAS ============================ */
function colTend(){ return {t:'Fallas por mes (1.ª → 2.ª mitad)',
  ayuda:'Fallas por mes en la primera y en la segunda mitad de los meses completos del '+
        'filtro. «Empeora» o «Mejora» solo si la diferencia es mayor que la variación normal.',
  k:function(f){ var t=(f._g||f).tend; return t ? {SUBE:0, BAJA:1, IGUAL:2}[t.estado] : 3; },
  csvk:function(f){ var t=(f._g||f).tend;
    return t ? evidencia(t)+' ('+TNOM[t.estado]+')' : 'pocos datos'; },
  html:function(f){ var t=(f._g||f).tend;
    if(!t) return '<span class="suave">pocos datos</span>';
    return '<span class="num">'+evidencia(t)+'</span>'+
      (t.estado==='IGUAL' ? '' : ' '+badgeT(t)); }}; }
function colEquipo(){ return {t:'Equipo', cls:'b',
  k:function(f){ return actFull(f.maq, f.mod); }}; }
function vMaquinas(rows){
  var maqs = porHoras(porMaquina(rows)), tot = sumaH(rows) || 1;
  var filas = maqs.map(function(g){
    var peor = null, pv = -1;
    Object.keys(g.mods).forEach(function(m){ if(g.mods[m]>pv){ pv=g.mods[m]; peor=+m; } });
    return {maq:D.dic.maquina[g.maq], n:g.n, h:g.h, pct:g.h/tot*100, mttr:g.mttr, hd:g.h/DIAS,
            peor: peor==null?'—':modLabel(peor), peorCod: peor==null?'—':modCod(peor),
            ppeor: g.h>0?pv/g.h*100:0, _g:g}; });
  tabla(el('m-tabla'), 'maquinas', {
    titulo:'Una fila por máquina. Clic para filtrar el reporte por ella.', archivo:'maquinas',
    per:12, buscar:false,
    cols:[
      {t:'Máquina', k:'maq', cls:'b'},
      {t:'Horas de avería', k:'h', tipo:'num', d:1, bar:true, fmt:function(v){return fh(v);}},
      {t:'% del total', k:'pct', tipo:'num', d:1, fmt:function(v){return fpc(v,1);}},
      {t:'Horas por día', k:'hd', tipo:'num', d:2, fmt:function(v){return fh(v);}},
      {t:'Averías', k:'n', tipo:'num', d:0},
      {t:'Min./avería', k:'mttr', tipo:'num', d:1, fmt:function(v){return fmin(v);},
       csvk:function(f){ return (f.mttr*60).toFixed(1); }},
      {t:'Módulo que más pesa', k:'peor',
       html:function(f){ return '<span title="'+esc(f.peor)+'"><b>'+esc(f.peorCod)+'</b> '+
         '<span class="suave">'+fpc(f.ppeor,0)+'</span></span>'; }},
      colTend()
    ], filas: filas,
    rowKey:function(f){ return String(f._g.maq); },
    onRow:function(k){ S.sel.maquina = [+k]; render(); }
  });
  gMensual(el('m-mens'), rows, el('m-mensleg'));
  gMatriz(el('m-matriz'), rows);
  gPareto(el('m-pareto'), porActivo(rows), {top: esMovil()?10:15});
}

/* ============================ PRIORIDADES ============================ */
function vPrioridades(rows){
  var acts = porActivo(rows), res = cuadrantes(acts), tot = sumaH(rows) || 1;
  el('p-grupos').innerHTML = QS.map(function(q){
    var g = res.todos.filter(function(x){return x.q===q;});
    var h = suma(g.map(function(x){return x.h;}));
    return '<div class="qcard" style="border-top-color:'+QCOL[q]+'"><div class="v">'+
      fpc(h/tot*100,0)+'<i>de las horas</i></div><div class="k">'+QNOM[q]+' · '+
      fnum(g.length,0)+' '+plural(g.length,'equipo','equipos')+'</div><div class="d">'+
      QSUB[q]+'</div></div>'; }).join('');
  gJK(el('p-jk'), acts, res, el('p-jkleg'));

  var sube = porHoras(acts.filter(function(g){ return g.tend && g.tend.estado==='SUBE'; }));
  var baja = porHoras(acts.filter(function(g){ return g.tend && g.tend.estado==='BAJA'; }));
  var conT = acts.filter(function(g){ return g.tend; }).length;
  var li = function(g){ return '<li data-k="'+g.maq+':'+g.mod+'"><span>'+esc(g.full)+
    '</span><span class="n">de <b>'+tasa(g.tend.ini)+'</b> a <b>'+tasa(g.tend.fin)+
    '</b> fallas/mes · '+fh(g.h)+' h</span></li>'; };
  var lista = function(a, vacioTxt){ return a.length ? '<ul class="lista2">'+
    a.slice(0,10).map(li).join('')+'</ul>'+(a.length>10 ? '<p class="suave" '+
    'style="margin:8px 0 0">y '+(a.length-10)+' más en la tabla de abajo.</p>' : '')
    : '<p class="suave" style="margin:0">'+vacioTxt+'</p>'; };
  if(VM.length < M.minMeses){
    el('p-sube').innerHTML = '<p class="suave" style="margin:0">Con este rango hay '+VM.length+
      ' '+plural(VM.length,'mes completo','meses completos')+'; para medir la tendencia hacen falta '+
      M.minMeses+'.</p>';
    el('p-baja').innerHTML = '';
  } else {
    el('p-sube').innerHTML = lista(sube, 'Ningún equipo empeora de forma clara.');
    el('p-baja').innerHTML = lista(baja, 'Ningún equipo mejora de forma clara.');
  }
  var h1 = VM.length ? Math.floor(VM.length/2) : 0;
  el('p-tnota').textContent = VM.length >= M.minMeses ? 'Compara las fallas por mes de '+
    mesTxt(VM[0].k)+' a '+mesTxt(VM[h1-1].k)+' con las de '+mesTxt(VM[h1].k)+' a '+
    mesTxt(VM[VM.length-1].k)+'. Solo aparece un equipo si la diferencia es mayor que la '+
    'variación normal de un mes a otro. Se evaluaron los '+fnum(conT,0)+' equipos con al '+
    'menos '+M.minTend+' fallas; en los demás no hay datos suficientes.' : '';
  document.querySelectorAll('#p-sube li, #p-baja li').forEach(function(x){
    x.addEventListener('click', function(){ filtraActivo(x.getAttribute('data-k')); }); });

  var orden = {CRITICO:0, CRONICO:1, AGUDO:2, LEVE:3};
  tabla(el('p-tabla'), 'prioridades', {
    titulo:'Todos los equipos, ordenados por horas. Clic en una fila para filtrar.',
    archivo:'prioridades', per:15,
    cols:[
      colEquipo(),
      {t:'Grupo', k:function(f){return orden[f.q]+' '+QNOM[f.q];},
       html:function(f){ return badgeQ(f.q); }, csvk:function(f){ return QNOM[f.q]; }},
      {t:'Horas', k:'h', tipo:'num', d:1, bar:true, fmt:function(v){return fh(v);}},
      {t:REC.length ? 'Últimos '+REC.length+' meses' : 'Recientes', k:'hRec',
       tipo:'num', d:1, fmt:function(v){ return REC.length ? fh(v) : '—'; },
       ayuda:'Horas de avería en los últimos meses completos del filtro'},
      {t:'Fallas', k:'n', tipo:'num', d:0},
      {t:'Min./falla', k:'mttr', tipo:'num', d:1, fmt:function(v){return fmin(v);},
       csvk:function(f){ return (f.mttr*60).toFixed(1); }},
      colTend()
    ], filas: porHoras(acts),
    rowKey:function(f){ return f.maq+':'+f.mod; },
    onRow:filtraActivo
  });
}

/* ============================ FALLAS ============================ */
function vFallas(rows){
  var tot = sumaH(rows) || 1;
  gPareto(el('f-fam'), porDim(rows,'familia'), {top: esMovil()?8:12, modo:'parte', etq:240,
    onClick:function(k){ S.sel.familia=[+k]; render(); }});
  var g = {}, i, r, k;
  for(i=0;i<rows.length;i++){
    r = rows[i]; k = r[F.L3]+'|'+r[F.L4];
    if(!g[k]) g[k] = {l3:r[F.L3], l4:r[F.L4], n:0, h:0, maqs:{}};
    g[k].n++; g[k].h += r[F.H]; g[k].maqs[r[F.MAQ]] = 1; }
  var filas = Object.keys(g).map(function(x){ var o=g[x]; o.mttr=o.h/o.n;
    o.pct=o.h/tot*100; o.nm=Object.keys(o.maqs).length; return o; })
    .sort(function(a,b){return b.h-a.h;});
  tabla(el('f-sint'), 'sintomas', {
    titulo:'Sistema y detalle registrados, ordenados por horas', archivo:'sintomas', per:10,
    cols:[
      {t:'Sistema (L3)', k:function(f){return D.dic.l3[f.l3];}, cls:'b'},
      {t:'Detalle (L4)', k:function(f){return D.dic.l4[f.l4];},
       html:function(f){ return esc(D.dic.l4[f.l4])+
         (D.gen[f.l4] ? ' <span class="gen">genérico</span>' : ''); }},
      {t:'Horas', k:'h', tipo:'num', d:1, bar:true, fmt:function(v){return fh(v);}},
      {t:'% del total', k:'pct', tipo:'num', d:1, fmt:function(v){return fpc(v,1);}},
      {t:'Averías', k:'n', tipo:'num', d:0},
      {t:'Min. por avería', k:'mttr', tipo:'num', d:1, fmt:function(v){return fmin(v);},
       csvk:function(f){ return (f.mttr*60).toFixed(1); }},
      {t:'Máquinas', k:'nm', tipo:'num', d:0}
    ], filas: filas
  });
  var rep = {};
  rows.forEach(function(rr){
    if(SC.l4[rr[F.L4]]) return;
    var kk = rr[F.MAQ]+'|'+rr[F.MOD]+'|'+rr[F.L4];
    if(!rep[kk]) rep[kk] = {maq:rr[F.MAQ], mod:rr[F.MOD], l4:rr[F.L4], n:0, h:0};
    rep[kk].n++; rep[kk].h += rr[F.H]; });
  var lr = Object.keys(rep).map(function(x){return rep[x];})
             .filter(function(o){ return o.n >= 10; })
             .sort(function(a,b){ return b.n-a.n; });
  tabla(el('f-rep'), 'repetidos', {
    titulo:'El mismo detalle 10 veces o más en el mismo equipo',
    nota:'· candidatos a análisis de causa raíz', archivo:'repetidos', per:10,
    cols:[
      colEquipo(),
      {t:'Detalle (L4)', k:function(f){return D.dic.l4[f.l4];},
       html:function(f){ return esc(D.dic.l4[f.l4])+
         (D.gen[f.l4] ? ' <span class="gen">genérico</span>' : ''); }},
      {t:'Veces', k:'n', tipo:'num', d:0, bar:true},
      {t:'Horas', k:'h', tipo:'num', d:1, fmt:function(v){return fh(v);}},
      {t:'Min. por avería', k:function(f){return f.h/f.n;}, tipo:'num', d:1,
       fmt:function(v){return fmin(v);}, csvk:function(f){ return (f.h/f.n*60).toFixed(1); }}
    ], filas: lr,
    vacioT:'Ningún detalle se repite 10 veces o más aquí',
    vacio:'Con el filtro actual no hay repeticiones suficientes.',
    rowKey:function(f){ return f.maq+':'+f.mod; },
    onRow:filtraActivo
  });
  if(el('f-exp').open) pintaExplorador(rows);
}
function pintaExplorador(rows){
  tabla(el('f-eventos'), 'eventos', {
    titulo:'Cada fila es una avería (los registros de una misma parada ya están unidos)',
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
      {t:'Minutos', k:function(f){return f[F.H]*60;}, tipo:'num', d:1},
      {t:'Registros unidos', k:function(f){return f[F.N];}, tipo:'num', d:0,
       ayuda:'Cuántos registros del archivo forman esta avería'}
    ], filas: rows,
    vacioT:'Ninguna avería coincide con el filtro'
  });
}

/* ============================ orquestacion ============================ */
var VISTAS = {resumen:vResumen, maquinas:vMaquinas, prioridades:vPrioridades,
              fallas:vFallas, notas:function(){}};
var RENDER_OK = true;
function render(){
  calculaVentana();
  var rows = filtra(null);
  pintaChips(); pintaSelects(); pintaMeses(); pintaPills(); pintaEstado(rows); pintaMas();
  document.querySelectorAll('.vista').forEach(function(v){
    v.classList.toggle('on', v.getAttribute('data-v') === S.vista); });
  document.querySelectorAll('.tab').forEach(function(t){
    t.setAttribute('aria-selected', String(t.getAttribute('data-v') === S.vista)); });
  try { VISTAS[S.vista](rows); }
  catch(err){ if(RENDER_OK){ RENDER_OK = false;
    console.error('Error al dibujar la vista '+S.vista, err); } }
}
function cambiaVista(v){
  if(!VISTAS[v]) return;
  S.vista = v; tipHide(); render();
  var m = document.querySelector('main');
  if(m && window.scrollY > m.offsetTop) window.scrollTo(0, Math.max(0, m.offsetTop-60));
}
function init(){
  TIP = el('tip');
  D.t0ms = Date.parse(D.t0.replace(' ','T')+'Z');
  D.finms = Date.parse(D.fin.replace(' ','T')+'Z');
  marcaSinCod();
  D.cov.forEach(function(c){
    var k = c[0];
    COV[k] = {d:c[1], dm:c[2], a:Math.max(D.t0ms, mesMs(k)), b:Math.min(D.finms, mesMs(k+1))};
    MESKEYS.push(k); });
  D.rows.forEach(function(r){
    r[F.MK] = mesKeyMs(D.t0ms + r[F.T]*3600000);
    if(!COV[r[F.MK]]){ COV[r[F.MK]] = {d:1, dm:30, a:mesMs(r[F.MK]), b:mesMs(r[F.MK]+1)};
                       MESKEYS.push(r[F.MK]); } });
  MESKEYS.sort(function(a,b){ return a-b; });
  S.m0 = MESKEYS[0]; S.m1 = MESKEYS[MESKEYS.length-1];

  document.querySelectorAll('.tab').forEach(function(t){
    t.addEventListener('click', function(){ cambiaVista(t.getAttribute('data-v')); }); });
  var tq;
  el('fq').addEventListener('input', function(){
    var v = this.value.toLowerCase(); clearTimeout(tq);
    tq = setTimeout(function(){ S.q = v; render(); }, 240); });
  el('fm0').addEventListener('change', function(){
    S.m0 = +this.value; if(S.m1 < S.m0) S.m1 = S.m0; render(); });
  el('fm1').addEventListener('change', function(){
    S.m1 = +this.value; if(S.m0 > S.m1) S.m0 = S.m1; render(); });
  el('fmas').addEventListener('click', function(){ S.mas = !S.mas; pintaMas(); });
  el('f-exp').addEventListener('toggle', function(){
    if(this.open) pintaExplorador(filtra(null)); });
  var imp = el('imprimir');
  if(imp) imp.addEventListener('click', function(){ window.print(); });
  window.addEventListener('beforeprint', function(){
    var rows = filtra(null);
    Object.keys(VISTAS).forEach(function(v){ try { VISTAS[v](rows); } catch(e){} }); });
  document.addEventListener('click', function(){ if(ABIERTO){ ABIERTO=null; pintaSelects(); } });
  document.addEventListener('keydown', function(e){
    if(e.key==='Escape'){ if(ABIERTO){ ABIERTO=null; pintaSelects(); } tipHide(); } });
  var rt, w0 = window.innerWidth;
  window.addEventListener('resize', function(){
    if(window.innerWidth === w0) return;
    w0 = window.innerWidth; clearTimeout(rt); rt = setTimeout(render, 220); });
  presets();
  render();
}
if(document.readyState === 'loading')
  document.addEventListener('DOMContentLoaded', init);
else init();
})();
"""


# =============================================================================
# 7. DOCUMENTO
# =============================================================================
PESTANAS = [("resumen", "Resumen"), ("maquinas", "Máquinas"),
            ("prioridades", "Prioridades"), ("fallas", "Fallas"), ("notas", "Notas")]
MESES_L = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto",
           "setiembre", "octubre", "noviembre", "diciembre"]
QNOM = {"CRITICO": "Crítico", "CRONICO": "Crónico", "AGUDO": "Agudo", "LEVE": "Leve"}
TNOM = {"SUBE": "empeora", "BAJA": "mejora", "IGUAL": "sin cambio claro"}


def mes_txt(k):
    return f"{MESES_L[k % 12]} de {k // 12}"


def lee(titulo, cuerpo):
    return (f'<details class="lee"><summary>{titulo}</summary>'
            f'<div class="cuerpo">{cuerpo}</div></details>')


def notas_html(dq, info, n_av):
    """Pestaña Notas: calidad del dato, cómo se calcula y qué se dejó de lado."""
    cal = [
        f"<li><b>{dq['n_reg']:,} registros del archivo → {n_av:,} averías.</b> Se unieron "
        f"{info['unidos']:,} registros que eran parte de una misma parada: en el mismo equipo, "
        f"el siguiente empezó menos de {UNIR_MINUTOS} minutos después de que terminó el "
        f"anterior. De ellos, {info['continuas']:,} empezaban menos de un minuto después: "
        f"una parada partida en el cambio de turno o un rearranque que no prosperó.</li>",
        f"<li><b>{dq['pct_generico_h']:.0f} % de las horas tiene un detalle (L4) genérico</b>, "
        f"como «Regulación electrónica», «Falla general», «Ajuste electrónico» u «Otros "
        f"ajustes». Dice dónde se paró, pero no qué pieza falló, y eso limita el análisis "
        f"de causa raíz.</li>",
        f"<li><b>{dq['pct_sin_cod_h']:.0f} % de las horas figura como «Libre» o sin "
        f"clasificar</b> ({dq['horas_sin_cod']:,.0f} h)"
        + (f"; la máquina con más horas así es {dq['peor_cod'][0]} "
           f"({dq['peor_cod'][1]:.0f} % de sus horas)" if dq.get("peor_cod") else "")
        + ".</li>",
        f"<li><b>{dq['pct_cortas']:.0f} % de los registros dura menos de 2 minutos</b> y "
        f"suma solo el {dq['pct_cortas_h']:.1f} % de las horas. Inflan la cantidad de fallas, "
        f"pero casi no pesan en tiempo: por eso el reporte ordena por horas.</li>",
    ]
    if dq.get("unificados"):
        cal.append(f"<li>Se unificaron {dq['unificados']:,} variantes de escritura de "
                   f"Sistema y Detalle (mayúsculas, tildes y prefijos como «Elec -» o "
                   f"«Mech -») para que cuenten juntas. Se apaga con UNIFICAR_TEXTOS.</li>")
    for k, dd, dm in dq["cov"]:
        if dd < 0.5 * dm:
            cal.append(f"<li><b>{mes_txt(k).capitalize()} tiene solo {dd:.1f} días de "
                       f"datos.</b> Aparece en los gráficos, pero no entra en el cálculo de "
                       f"tendencias ni debe compararse con los demás meses.</li>")
    if dq.get("sugerido_inicio") is not None:
        cal.append(f"<li><b>El registro arranca irregular.</b> Los primeros meses tienen muy "
                   f"pocos eventos frente al resto; desde "
                   f"{dq['sugerido_inicio'].strftime('%m/%Y')} el registro es estable (botón "
                   f"«Registro estable» en la barra de filtros).</li>")

    return f"""
  <div class="vhead"><h2>Notas</h2>
    <p>Qué tan confiable es el dato, cómo se calcula cada número y qué se dejó fuera.</p></div>

  <div class="card">
    <h3>Calidad del dato</h3>
    <p class="lead">Todo el reporte depende de cómo se registró la parada en planta.</p>
    <ul class="simple">{''.join(cal)}</ul>
  </div>

  <div class="card">
    <h3>Cómo se calcula</h3>
    <div class="defs">
      <div><h4>Avería</h4><p>Una parada de un equipo. Los registros del mismo equipo
        separados por menos de {UNIR_MINUTOS} minutos (del fin de uno al inicio del
        siguiente) se cuentan como una sola avería; el sistema y el detalle se toman del
        registro más largo.</p></div>
      <div><h4>Equipo</h4><p>Un módulo de una máquina, por ejemplo PF4 · CS. Es el nivel al
        que tiene sentido asignar un responsable.</p></div>
      <div><h4>Horas de avería</h4><p>Suma de la duración de las averías. Es la medida con la
        que se ordena casi todo el reporte.</p></div>
      <div><h4>Minutos por avería (MTTR)</h4><p>Horas de avería ÷ cantidad de averías. Dice
        cuánto demora, en promedio, volver a producir.</p></div>
      <div><h4>Grupos: crítico, crónico, agudo y leve</h4><p>Método Jack-Knife de Knights
        (2001). Un equipo es <b>crónico</b> si falla más veces que el promedio de los equipos,
        <b>agudo</b> si cada falla dura más que el promedio de todas las fallas, y
        <b>crítico</b> si cumple las dos cosas.</p></div>
      <div><h4>Tendencia</h4><p>Compara las fallas por mes de la segunda mitad de los meses
        completos (los que tienen al menos la mitad de los días con datos) con las de la
        primera mitad. Solo dice «empeora» o «mejora» si la diferencia es mayor que la
        variación normal de un mes a otro de ese mismo equipo (razón de tasas cuasi-Poisson,
        95 %); un mes malo aislado no alcanza. Necesita al menos {MIN_FALLAS_TENDENCIA}
        fallas y {MIN_MESES_TENDENCIA} meses completos.</p></div>
      <div><h4>Dónde actuar primero</h4><p>Los {TOP_ACCIONES} equipos con más horas de avería
        en los últimos {MESES_RECIENTES} meses completos del filtro (o en todo el filtro si
        es más corto). Así no aparece arriba algo que ya se resolvió.</p></div>
      <div><h4>Detalle genérico</h4><p>Un detalle (L4) que describe la intervención y no la
        falla: regulación, calibración, ajuste electrónico, otros ajustes, falla general,
        reparación mecánica, cambio o reemplazo de componentes, libre.</p></div>
    </div>
  </div>

  <div class="card">
    <h3>Qué se dejó fuera respecto a la versión anterior, y por qué</h3>
    <ul class="simple">
      <li><b>Weibull por módulo.</b> Salía β menor que 1 («mortalidad infantil») en casi todos
        los equipos. Eso no indica repuestos o reparaciones de mala calidad: un módulo junta
        muchas piezas distintas y sus fallas llegan en rachas, y eso tiende a dar β menor
        que 1.
        Para usar Weibull en el plan de preventivo hace falta el tiempo de operación entre
        cambios de una misma pieza.</li>
      <li><b>Crow-AMSAA para la tendencia.</b> Supone que las fallas llegan de a una y al
        azar. Aquí llegan en rachas (un mes malo pesa mucho) y marcaba como «empeorando» o
        «mejorando» a muchos equipos sin un cambio sostenido. La tendencia actual toma en
        cuenta esa variación.</li>
      <li><b>Disponibilidad.</b> Se calculaba como 1 − horas de avería ÷ horas del calendario.
        Sin el tiempo programado, las paradas planificadas y las demás pérdidas, ese número
        no es la disponibilidad de la línea.</li>
      <li><b>Mediana en frecuencia contra duración.</b> Se usaba la mediana como límite, que no
        es el método de Knights que se citaba. Ahora se usan sus límites (promedios).</li>
    </ul>
  </div>

  <div class="card">
    <h3>Qué no se puede afirmar con esta data</h3>
    <ul class="simple">
      <li>La disponibilidad real: falta el tiempo programado de producción.</li>
      <li>El costo de la parada: no hay valorización económica.</li>
      <li>La causa raíz: el registro anota dónde se paró y un síntoma, no la causa.</li>
      <li>El efecto del preventivo: habría que cruzar con las órdenes ejecutadas.</li>
    </ul>
  </div>"""


def construir_html(pl, dq, info):
    m = pl["meta"]
    js = JS.replace("__DATA__", json.dumps(pl, ensure_ascii=False,
                                           separators=(",", ":")).replace("<", "\\u003c"))
    tabs = "".join(
        f'<button class="tab" role="tab" data-v="{k}" aria-selected="'
        f'{"true" if k == "resumen" else "false"}">{t}</button>'
        for k, t in PESTANAS)
    desde, hasta = dq["ini"], dq["fin"] - pd.Timedelta(seconds=1)
    trunc = ("" if not m["truncado"] else
             f'<div class="nota"><p>El reporte contiene las {m["nEmbed"]:,} averías de mayor '
             f'duración de las {m["nFinal"]:,} del periodo. Sube MAX_FILAS_TABLA en el script '
             f'si necesitas todas.</p></div>')

    return f"""<!DOCTYPE html>
<html lang="es"><head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1,viewport-fit=cover">
<title>{TITULO} · {PLANTA}</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@500;600&family=IBM+Plex+Sans:wght@400;500;600;700&display=swap" rel="stylesheet">
<style>{CSS}</style></head>
<body>

<header class="rail">
  <div class="wrap">
    <div class="marca">
      <p class="planta">{PLANTA}</p>
      <h1>{TITULO}</h1>
      <p class="sub">{SUBTITULO}</p>
    </div>
    <div class="railmeta">
      Del {desde:%d/%m/%Y} al {hasta:%d/%m/%Y}<br>
      <b>{m['nFinal']:,}</b> averías · <b>{m['horasTotal']:,.0f} h</b><br>
      Emitido el {m['emitido']}
      <br><button class="imp" id="imprimir">Imprimir o guardar en PDF</button>
    </div>
  </div>
</header>

<nav class="tabs" role="tablist" aria-label="Secciones del reporte">
  <div class="wrap">{tabs}<span class="filtroact" id="tfiltro"></span></div>
</nav>

<div class="filtros">
  <div class="wrap">
    <div class="frow">
      <span class="fetq">Máquinas</span>
      <div class="frow" id="fchips" style="gap:6px"></div>
    </div>
    <div class="frow">
      <span class="fetq">Meses</span>
      <select id="fm0" aria-label="Mes inicial"></select>
      <span class="suave" style="font-size:13.4px">a</span>
      <select id="fm1" aria-label="Mes final"></select>
      <div class="frow" id="fpresets" style="gap:6px"></div>
      <button class="btn g" id="fmas" aria-expanded="false" aria-controls="fmasbox">Más filtros</button>
      <span class="estado" id="festado"></span>
    </div>
    <div class="frow oculto" id="fmasbox">
      <div class="frow" id="fsel" style="gap:8px"></div>
      <input type="search" id="fq" placeholder="Buscar síntoma, módulo, sistema"
             style="min-width:230px" aria-label="Buscar texto libre">
    </div>
    <div class="frow oculto" id="fpills" style="gap:6px"></div>
  </div>
</div>

<main><div class="wrap">

<!-- ============ RESUMEN ============ -->
<section class="vista on" data-v="resumen" role="tabpanel">
  <div class="titular" id="r-titular"></div>
  <div class="kpis" id="r-kpis"></div>

  <div class="cols c2">
    <div class="card">
      <h3>Dónde se pierde</h3>
      <p class="lead">Horas de avería por máquina, separadas por tipo. Clic para filtrar.</p>
      <div class="chart" id="r-maq"></div>
      <div class="leg" id="r-maqleg"></div>
    </div>
    <div class="card">
      <h3>Qué falla</h3>
      <p class="lead">Familias de falla por horas, sumando todas las máquinas. Clic para filtrar.</p>
      <div class="chart" id="r-fam"></div>
    </div>
  </div>

  <div class="card">
    <h3>Dónde actuar primero</h3>
    <p class="lead" id="r-accsub"></p>
    <div id="r-acciones"></div>
    {lee("Cómo se eligen y qué significa cada etiqueta",
        "<p>Se ordenan por las horas perdidas en los últimos meses completos, no en todo el "
        "periodo, para no poner arriba algo que ya se resolvió.</p>"
        "<p><b>Crítico</b>: falla más seguido que el promedio y cada parada dura más que el "
        "promedio. <b>Crónico</b>: falla seguido pero se arregla rápido; el problema es que "
        "vuelve a fallar. <b>Agudo</b>: falla poco, pero cada parada es larga; el problema es "
        "cuánto demora la reparación. <b>Empeora</b> o <b>Mejora</b> aparece solo cuando el "
        "cambio en la cantidad de fallas se sostiene mes a mes.</p>"
        "<p>«Lo que más se repite» es el detalle (L4) registrado con más frecuencia en ese "
        "equipo. Si es genérico, el registro no alcanza para saber qué pieza falló.</p>")}
  </div>

  <div id="r-avisos"></div>
</section>

<!-- ============ MAQUINAS ============ -->
<section class="vista" data-v="maquinas" role="tabpanel">
  <div class="vhead"><h2>Máquinas</h2>
    <p>Cuánto pierde cada máquina, cómo evoluciona y qué módulo la está parando. Todo
    responde a los filtros de arriba.</p></div>

  <div class="card"><div id="m-tabla"></div></div>

  <div class="card">
    <h3>Mes a mes</h3>
    <p class="lead">Horas de avería de cada mes, divididas por máquina.</p>
    <div class="chart" id="m-mens"></div>
    <div class="leg" id="m-mensleg"></div>
  </div>

  <div class="card">
    <h3>Qué módulo falla en cada máquina</h3>
    <p class="lead">Horas por cruce. Una fila oscura es una máquina con problemas; una columna
    oscura, un módulo que falla en toda la planta.</p>
    <div class="chart" id="m-matriz"></div>
    {lee("Cómo usar esta matriz",
        "<p>Sirve para decidir si un problema es local o de toda la planta. Si el color fuerte "
        "se reparte a lo largo de una columna, el módulo falla en varias líneas y la solución "
        "probablemente es una sola (diseño, repuesto o procedimiento). Si se concentra en una "
        "celda, es un tema de esa máquina.</p>")}
  </div>

  <div class="card">
    <h3>Los equipos que más horas pierden</h3>
    <p class="lead">Combinaciones de máquina y módulo ordenadas por horas, con el acumulado.</p>
    <div class="chart" id="m-pareto"></div>
  </div>
</section>

<!-- ============ PRIORIDADES ============ -->
<section class="vista" data-v="prioridades" role="tabpanel">
  <div class="vhead"><h2>Prioridades</h2>
    <p>No todas las averías se resuelven igual: las que se repiten piden causa raíz y las que
    paran horas piden reparar más rápido. Y conviene saber cuáles van a peor.</p></div>

  <div class="cols c4" id="p-grupos" style="margin-bottom:18px"></div>

  <div class="card">
    <h3>Frecuencia contra duración</h3>
    <p class="lead">Cada círculo es un equipo con al menos {MIN_FALLAS_GRAFICO} fallas (los
    demás cuentan en los totales de arriba). A la derecha falla más seguido, arriba cada falla
    dura más, y el tamaño son las horas.</p>
    <div class="chart" id="p-jk"></div>
    <div class="leg" id="p-jkleg"></div>
    {lee("Cómo se lee este gráfico",
        "<p>Las líneas grises marcan el promedio de fallas por equipo y la duración promedio de "
        "una falla (método de Knights). Arriba a la derecha están los críticos: fallan más "
        "que el promedio y además tardan más en volver.</p>"
        "<p>Las diagonales punteadas unen puntos con las mismas horas totales: dos equipos sobre "
        "la misma diagonal pierden lo mismo, aunque uno falle diez veces por una hora y el otro "
        "una vez por diez horas. La diferencia está en cómo se atacan.</p>")}
  </div>

  <div class="cols c2">
    <div class="card">
      <h3>Empeoran de forma clara</h3>
      <p class="lead">Fallas por mes en la primera y en la segunda mitad del rango.</p>
      <div id="p-sube"></div>
    </div>
    <div class="card">
      <h3>Mejoran de forma clara</h3>
      <p class="lead">Conviene saber qué se hizo para repetirlo en otros equipos.</p>
      <div id="p-baja"></div>
    </div>
  </div>
  <p class="suave" id="p-tnota" style="font-size:13.4px;margin:-4px 0 18px;max-width:100ch"></p>

  <div class="card"><div id="p-tabla"></div></div>
</section>

<!-- ============ FALLAS ============ -->
<section class="vista" data-v="fallas" role="tabpanel">
  <div class="vhead"><h2>Fallas</h2>
    <p>Qué se rompe, según lo que se registró. Es el insumo para el análisis de causa raíz.</p></div>
  {trunc}
  <div class="card">
    <h3>Familias de falla</h3>
    <p class="lead">Agrupación de los sistemas (L3) por horas. Clic en una barra para filtrar.</p>
    <div class="chart" id="f-fam"></div>
  </div>

  <div class="card">
    <h3>Qué se registró</h3>
    <p class="lead">Sistema (L3) y detalle (L4) ordenados por horas. «Genérico» marca los
    detalles que dicen que hubo que intervenir, pero no qué falló.</p>
    <div id="f-sint"></div>
  </div>

  <div class="card">
    <h3>Lo que se repite en el mismo equipo</h3>
    <p class="lead">Cuando el mismo detalle vuelve decenas de veces sobre el mismo equipo, la
    reparación está tratando el efecto y no la causa. Clic en una fila para filtrar.</p>
    <div id="f-rep"></div>
  </div>

  <div class="card">
    <details class="bloque" id="f-exp">
      <summary>Ver las averías una por una <span>con los filtros de arriba; se pueden
      descargar en CSV</span></summary>
      <div id="f-eventos"></div>
    </details>
  </div>
</section>

<!-- ============ NOTAS ============ -->
<section class="vista" data-v="notas" role="tabpanel">
{notas_html(dq, info, m['nFinal'])}
</section>

</div></main>

<footer><div class="wrap">
  <p><b>{TITULO}</b> · {PLANTA}. Del {desde:%d/%m/%Y} al {hasta:%d/%m/%Y}.
  Emitido el {m['emitido']}.</p>
  <p>Frecuencia contra duración: Jack-Knife de Knights (2001). Tendencia: fallas por mes de
  la segunda mitad contra la primera, con la variación real entre meses (cuasi-Poisson,
  95 %). Los registros del mismo equipo separados por menos de {UNIR_MINUTOS} minutos se
  cuentan como una sola avería.</p>
</div></footer>

<div class="tip" id="tip"></div>
<script>{js}</script>
</body></html>"""


# =============================================================================
# 8. MAIN
# =============================================================================
def main():
    ruta = buscar_archivo(sys.argv[1] if len(sys.argv) > 1 else ARCHIVO)
    print("=" * 70)
    print("  ANALISIS DE CONFIABILIDAD DE AVERIAS  ·  v5")
    print("=" * 70)
    if not ruta or not os.path.exists(ruta):
        print(f"\n  ERROR: no encuentro el archivo:\n  {ruta}")
        print("  Ajusta ARCHIVO al inicio del script o pásalo como argumento.\n")
        return
    print(f"\n  Archivo: {ruta}")

    df = cargar(ruta)
    unificados = df.attrs.get("unificados", 0)
    print(f"  [1/5] {len(df):,} registros · {df['Maquina'].nunique()} máquinas · "
          f"{df['Modulo'].nunique()} módulos")
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
    print(f"  [2/5] {len(av):,} averías (se unieron {info['unidos']:,} registros que eran "
          f"parte de una misma parada)")

    carpeta = os.path.dirname(os.path.abspath(ruta))
    dq["cobertura"] = av["Modulo"].str.upper().isin(nombres.keys()).mean() * 100
    faltan = exportar_pendientes(av, nombres, carpeta)
    print(f"  [3/5] Calidad: {dq['pct_generico_h']:.0f} % de las horas con detalle genérico"
          f" · {dq['pct_sin_cod_h']:.0f} % sin codificar"
          + (f" · {faltan} módulos sin nombre -> modulos_pendientes.csv" if faltan else ""))

    pl = payload(av, nombres, dq, info)
    salida = SALIDA_HTML or os.path.join(carpeta, "reporte_confiabilidad.html")
    try:
        with open(salida, "w", encoding="utf-8") as f:
            f.write(construir_html(pl, dq, info))
    except PermissionError:
        base, ext = os.path.splitext(salida)
        salida = f"{base}_{datetime.now():%Y%m%d_%H%M%S}{ext}"
        with open(salida, "w", encoding="utf-8") as f:
            f.write(construir_html(pl, dq, info))
    print(f"  [4/5] {salida} ({os.path.getsize(salida) / 1024:,.0f} KB)")

    g, rec = analisis_activos(av, dq["cov"])
    est = {"SUBE": 0, "BAJA": 0, "IGUAL": 0, None: 0}
    for t in g["tend"]:
        est[t["estado"] if t else None] += 1
    print(f"  [5/5] Tendencia en {len(meses_completos(dq['cov']))} meses completos: "
          f"empeoran {est['SUBE']} · mejoran {est['BAJA']} · sin cambio claro "
          f"{est['IGUAL']} · pocos datos {est[None]}")

    print("\n" + "=" * 70)
    print(f"  {av['Horas'].sum():,.0f} h de avería · {len(av):,} averías · "
          f"{len(g)} equipos (máquina · módulo)")
    orden = g[g["mod"] != SIN].sort_values(["h_rec", "h"] if rec else ["h"], ascending=False)
    print("  Dónde actuar primero" +
          (f" (horas de {mes_txt(rec[0])} a {mes_txt(rec[-1])}):" if rec else ":"))
    for i, (a, r) in enumerate(orden.head(TOP_ACCIONES).iterrows(), 1):
        t = r["tend"]
        print(f"    {i}. {a:<16} {(r['h_rec'] if rec else r['h']):7.1f} h · "
              f"{QNOM[r['q']]:<8} · {TNOM[t['estado']] if t else 'pocos datos'}")
    print("=" * 70)


if __name__ == "__main__":
    main()
