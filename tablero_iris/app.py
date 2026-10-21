# -*- coding: utf-8 -*-
"""Línea 195 - Panel de seguimiento de campañas SMS / Mailing.

Lee el Excel de seguimiento (hojas "Mailing" y "SMS"), lo limpia con etl.py
y expone una API que alimenta un dashboard responsive construido con
Flask + Chart.js.

Cómo se actualiza la información:
  1. Subiendo un Excel nuevo desde el botón "Actualizar datos" (reemplaza el actual).
  2. Editando y guardando el Excel en disco: el tablero detecta el cambio solo.

Variables de entorno (ver .env.example):
  ARCHIVO_EXCEL     ruta del Excel fuente (por defecto datos/datos.xlsx)
  IRIS_USUARIO      usuario de acceso al panel (por defecto SGAM195)
  IRIS_PASSWORD     contraseña de acceso al panel (si se deja vacía, no pide login)
  SECRET_KEY        clave para firmar la sesión (usa una fija en producción)
  FLASK_DEBUG       "1" para modo debug (nunca usar en producción)
  LOG_LEVEL         INFO, DEBUG, WARNING... (por defecto INFO)
"""
import base64
import logging
import os
import secrets
import threading
import uuid
from datetime import datetime
from functools import wraps
from io import BytesIO
from pathlib import Path

import numpy as np
import pandas as pd
from dotenv import load_dotenv
from flask import (Flask, jsonify, redirect, render_template, request,
                    send_file, session, url_for)

import etl

load_dotenv()

BASE = Path(__file__).resolve().parent
ARCHIVO = Path(os.environ.get("ARCHIVO_EXCEL", BASE / "datos" / "datos.xlsx")).resolve()
EXTENSIONES = {".xlsx", ".xlsm"}
MAX_FILAS_TABLA_GENERICA = 200
MAX_CATEGORIAS_GENERICO = 12
TOP_ENTIDADES = 10

IRIS_USUARIO = os.environ.get("IRIS_USUARIO", "SGAM195").strip()
IRIS_PASSWORD = os.environ.get("IRIS_PASSWORD", "Linea195").strip()
SECRET_KEY = os.environ.get("SECRET_KEY", "").strip() or os.urandom(32)
LOG_LEVEL = os.environ.get("LOG_LEVEL", "INFO").upper()
DEBUG = os.environ.get("FLASK_DEBUG", "0") == "1"

logging.basicConfig(
    level=LOG_LEVEL,
    format="%(asctime)s %(levelname)s %(name)s: %(message)s",
)
log = logging.getLogger("iris")

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 25 * 1024 * 1024  # 25 MB
app.config["JSON_AS_ASCII"] = False
app.secret_key = SECRET_KEY
app.config["SESSION_COOKIE_HTTPONLY"] = True
app.config["SESSION_COOKIE_SAMESITE"] = "Lax"

ARCHIVO.parent.mkdir(parents=True, exist_ok=True)

_candado = threading.Lock()
_cache = {"clave": None, "modo": None, "df": None, "df_camp": None, "hojas": []}


# ---------------------------------------------------------------- autenticación ---
def requiere_login(vista):
    """Si IRIS_PASSWORD está configurada, exige sesión iniciada.
    Sin esa variable, el panel queda abierto (útil para desarrollo local)."""
    @wraps(vista)
    def envoltura(*args, **kwargs):
        if not IRIS_PASSWORD:
            return vista(*args, **kwargs)
        if session.get("autenticado"):
            return vista(*args, **kwargs)
        if request.path.startswith("/api/"):
            return jsonify(error="Sesión no iniciada."), 401
        return redirect(url_for("login", siguiente=request.path))
    return envoltura


@app.get("/login")
def login():
    if not IRIS_PASSWORD or session.get("autenticado"):
        return redirect(url_for("inicio"))
    return render_template("login.html", error=None)


@app.post("/login")
def login_post():
    usuario = request.form.get("usuario", "").strip()
    clave = request.form.get("password", "")
    usuario_ok = secrets.compare_digest(usuario, IRIS_USUARIO)
    clave_ok = bool(IRIS_PASSWORD) and secrets.compare_digest(clave, IRIS_PASSWORD)
    if usuario_ok and clave_ok:
        session["autenticado"] = True
        session["usuario"] = usuario
        session.permanent = True
        log.info("Inicio de sesión correcto (%s) desde %s", usuario, request.remote_addr)
        siguiente = request.args.get("siguiente") or url_for("inicio")
        return redirect(siguiente)
    log.warning("Intento de inicio de sesión fallido (usuario=%r) desde %s", usuario, request.remote_addr)
    return render_template("login.html", error="Usuario o contraseña incorrectos."), 401


@app.get("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


# ---------------------------------------------------------------- utilidades ---
def marca_archivo():
    return str(ARCHIVO.stat().st_mtime_ns) if ARCHIVO.exists() else None


def numero(valor):
    if valor is None:
        return None
    try:
        if pd.isna(valor):
            return None
    except (TypeError, ValueError):
        pass
    return float(valor)


def entero(valor):
    n = numero(valor)
    return int(round(n)) if n is not None else None


def fecha_iso(valor):
    if valor is None or pd.isna(valor):
        return None
    return pd.Timestamp(valor).strftime("%Y-%m-%d")


def celda(valor):
    if valor is None or (not isinstance(valor, str) and pd.isna(valor)):
        return None
    if isinstance(valor, (pd.Timestamp, datetime)):
        return valor.strftime("%Y-%m-%d")
    if isinstance(valor, np.integer):
        return int(valor)
    if isinstance(valor, np.floating):
        return None if np.isnan(valor) else float(valor)
    if isinstance(valor, np.bool_):
        return bool(valor)
    return valor


# ------------------------------------------------------------------- carga ---
def _agregar_por_campana(df):
    """Cada fila de envío del Excel es una campaña independiente: ya NO se
    agrupan varias filas que compartan 'Consecutivo' en una sola campaña
    (antes se trataban como distintas 'olas' de un mismo envío y se sumaban
    sus resultados). Por eso este paso ya no agrega/suma nada: solo prepara
    una copia de `df` con un `campana_id` único por fila (para que los
    conteos de "campañas" cuenten exactamente una por fila) y normaliza
    entidad/entidades igual que antes."""
    grp = df.reset_index(drop=True).copy()
    grp["campana_id"] = grp.index.astype(str)
    grp["entidad"] = grp["entidad"].fillna("Sin especificar").replace("", "Sin especificar")
    grp["entidades"] = grp["entidades"].apply(
        lambda l: l if isinstance(l, list) and l else ["Sin especificar"]
    )
    return grp


def _todas_las_entidades(df_camp):
    """Lista ordenada de entidades individuales presentes en df_camp,
    "desempacando" las campañas con varias entidades para que cada una
    aparezca como opción de filtro independiente."""
    vistas = set()
    for lista in df_camp["entidades"]:
        if isinstance(lista, list):
            vistas.update(lista)
    return sorted(vistas)


def _cargar():
    """Carga (con caché) el Excel actual. Devuelve dict con modo y datos."""
    clave = marca_archivo()
    with _candado:
        if _cache["clave"] == clave and clave is not None:
            return _cache

        if not ARCHIVO.exists():
            log.warning("El archivo fuente no existe: %s", ARCHIVO)
            _cache.update(clave=clave, modo=None, df=None, df_camp=None, hojas=[])
            return _cache

        if etl.es_formato_campanas(ARCHIVO):
            df, hojas = etl.cargar_campanas(ARCHIVO)
            if df is not None and not df.empty:
                df_camp = _agregar_por_campana(df)
                log.info("Datos cargados en modo 'campanas': %d envíos, %d campañas", len(df), len(df_camp))
                _cache.update(clave=clave, modo="campanas", df=df, df_camp=df_camp, hojas=hojas)
                return _cache

        # modo genérico de respaldo para cualquier otro Excel
        with pd.ExcelFile(ARCHIVO) as xls:
            hojas = list(xls.sheet_names)
            primera = xls.sheet_names[0]
            gdf = pd.read_excel(xls, sheet_name=primera)
        gdf = gdf.dropna(how="all").dropna(axis=1, how="all")
        gdf.columns = [str(c).strip() for c in gdf.columns]
        log.info("Datos cargados en modo 'generico': %d filas", len(gdf))
        _cache.update(clave=clave, modo="generico", df=gdf.reset_index(drop=True), df_camp=None, hojas=hojas)
        return _cache


def _valores_seleccionados(args, nombre, sentinela):
    """Lee los valores elegidos para un filtro de selección múltiple
    (checkboxes): varios parámetros de query con el mismo nombre
    (?entidad=A&entidad=B), descartando vacíos y el valor "Todas"/"Todos"
    de respaldo. Una lista vacía significa "sin filtro" (se muestra todo),
    igual que antes se interpretaba "Todas"/"Todos"."""
    return [v for v in args.getlist(nombre) if v and v != sentinela]


def _filtrar(df, args):
    if df is None or df.empty:
        return df
    out = df
    tipo = args.get("tipo")
    if tipo and tipo != "Todos":
        out = out[out["tipo"] == tipo]
    entidades_sel = _valores_seleccionados(args, "entidad", "Todas")
    if entidades_sel:
        conjunto = set(entidades_sel)
        out = out[out["entidades"].apply(
            lambda l: isinstance(l, list) and bool(conjunto.intersection(l))
        )]
    meses_sel = _valores_seleccionados(args, "mes", "Todos")
    if meses_sel:
        out = out[out["mes"].isin(meses_sel)]
    anios_sel = _valores_seleccionados(args, "anio", "Todos")
    if anios_sel:
        out = out[out["anio_txt"].isin(anios_sel)]
    buscar = (args.get("buscar") or "").strip().lower()
    if buscar:
        mask = (
            out["campana"].str.lower().str.contains(buscar, na=False)
            | out["objetivo"].str.lower().str.contains(buscar, na=False)
            | out["caracterizacion"].str.lower().str.contains(buscar, na=False)
            | out["entidad"].str.lower().str.contains(buscar, na=False)
        )
        out = out[mask]
    return out


# ------------------------------------------------------------------- rutas ---
@app.get("/salud")
def salud():
    """Chequeo de salud simple para Docker/orquestadores. No requiere sesión
    a propósito: solo confirma que el proceso Flask está vivo y respondiendo,
    sin exponer datos ni depender de que ya haya un Excel cargado."""
    return jsonify(estado="ok"), 200


@app.get("/")
@requiere_login
def inicio():
    return render_template("index.html", tiene_login=bool(IRIS_PASSWORD))


@app.get("/api/estado")
@requiere_login
def estado():
    return jsonify(marca=marca_archivo())


@app.get("/api/dashboard")
@requiere_login
def dashboard():
    estado_cache = _cargar()
    if not ARCHIVO.exists() or estado_cache["modo"] is None:
        return jsonify(hay_datos=False)

    meta = {
        "nombre": ARCHIVO.name,
        "hojas": estado_cache["hojas"],
        "marca": estado_cache["clave"],
        "actualizado": datetime.fromtimestamp(ARCHIVO.stat().st_mtime).astimezone().isoformat(),
    }

    if estado_cache["modo"] == "generico":
        return _dashboard_generico(estado_cache["df"], meta)
    return _dashboard_campanas(estado_cache["df"], estado_cache["df_camp"], meta)


def _dashboard_campanas(df, df_camp, meta):
    args = request.args
    datos = _calcular_dashboard_campanas(df, df_camp, args)

    if datos is None:
        return jsonify(
            hay_datos=True, modo="campanas", meta=meta,
            filtros_disponibles=_filtros_disponibles(df_camp),
            sin_resultados=True,
        )

    return jsonify(
        hay_datos=True, modo="campanas", meta=meta,
        filtros_disponibles=_filtros_disponibles(df_camp),
        seleccion=datos["seleccion"],
        kpis=datos["kpis"],
        comparativo_canal=datos["comparativo_canal"],
        evolucion=datos["evolucion"],
        evolucion_periodos=datos["evolucion_periodos"],
        top_entidades=datos["top_entidades"],
        efectividad_entidades=datos["efectividad_entidades"],
        resultado_envios=datos["resultado_envios"],
        ranking=datos["ranking"],
        tabla=datos["tabla"],
    )


_MESES_ABREV = {1: "ene", 2: "feb", 3: "mar", 4: "abr", 5: "may", 6: "jun",
                 7: "jul", 8: "ago", 9: "sept", 10: "oct", 11: "nov", 12: "dic"}


def _etiqueta_semana(fecha):
    """'8 de sept 2025' para el eje X del gráfico semanal. Siempre incluye
    el año: los datos reales de Línea 195 llegan a cubrir más de un año, y
    sin el año fechas como '16 de sept' quedarían ambiguas entre años."""
    return f"{fecha.day} de {_MESES_ABREV[fecha.month]} {fecha.year}"


def _series_por_periodo(base, orden, etiqueta):
    """Agrupa filas de envío (`base`, a nivel de ola/envío individual, ya
    filtradas) por un periodo de tiempo -una semana o un mes- y por canal
    (tipo), sumando 3 métricas seleccionables: mensajes efectivos, número
    de envíos y valor. `orden` es la clave por la que se ordena
    cronológicamente cada periodo; `etiqueta` es el texto que se muestra en
    el eje X. Devuelve {etiquetas: [...], series: {canal: {metrica: [...]}}}."""
    tmp = base.assign(_orden=orden, _etiqueta=etiqueta)
    periodos = (
        tmp[["_orden", "_etiqueta"]]
        .drop_duplicates()
        .sort_values("_orden")
    )
    etiquetas = periodos["_etiqueta"].tolist()
    indice = {o: i for i, o in enumerate(periodos["_orden"].tolist())}

    series = {}
    for tipo, g in tmp.groupby("tipo"):
        efectivos = [0] * len(etiquetas)
        envios = [0] * len(etiquetas)
        valor = [0.0] * len(etiquetas)
        for orden_val, gg in g.groupby("_orden"):
            i = indice[orden_val]
            efectivos[i] = entero(numero(gg["efectivos"].sum()) or 0)
            envios[i] = entero(len(gg))
            valor[i] = numero(gg["valor"].sum()) or 0
        series[str(tipo)] = {"efectivos": efectivos, "envios": envios, "valor": valor}

    return {"etiquetas": etiquetas, "series": series}


def _evolucion_periodos(df_f):
    """Series temporales de envíos por canal (SMS/Mailing) para el gráfico
    'Mensajes efectivos por semana': agrupadas por semana y por mes a la
    vez -el panel cambia entre ellas al instante, sin pedir datos de nuevo
    al servidor-, con 3 métricas seleccionables (mensajes efectivos,
    envíos, valor)."""
    vacio = {"semana": {"etiquetas": [], "series": {}}, "mes": {"etiquetas": [], "series": {}}}
    if df_f is None or df_f.empty:
        return vacio
    base = df_f.dropna(subset=["fecha_evento"])
    if base.empty:
        return vacio

    semana_inicio = base["fecha_evento"].dt.to_period("W-SUN").dt.start_time
    return {
        "semana": _series_por_periodo(base, semana_inicio, semana_inicio.apply(_etiqueta_semana)),
        "mes": _series_por_periodo(base, base["periodo_orden"], base["periodo"]),
    }


def _efectividad_por_entidad(df_f):
    """Ranking de entidades por volumen de envíos (igual criterio que
    `top_entidades`, pero contando olas/envíos individuales en vez de
    campañas) con su efectividad, para la lista 'Efectividad por entidad'."""
    if df_f is None or df_f.empty:
        return []
    entidades_seguras = df_f["entidades"].apply(
        lambda l: l if isinstance(l, list) and l else ["Sin especificar"]
    )
    explotado = df_f.assign(entidades=entidades_seguras).explode("entidades", ignore_index=True)
    resultado = []
    for entidad, g in explotado.groupby("entidades"):
        ev = numero(g["destinatarios_evaluables"].sum()) or 0
        efe = numero(g["efectivos"].sum()) or 0
        resultado.append({
            "entidad": entidad,
            "envios": entero(len(g)),
            "efectividad": round(efe / ev * 100, 2) if ev else None,
        })
    resultado.sort(key=lambda x: x["envios"], reverse=True)
    return resultado[:TOP_ENTIDADES]


def _calcular_dashboard_campanas(df, df_camp, args):
    """Calcula todas las secciones del dashboard (KPIs, comparativo por
    canal, evolución, entidades, resultados de envío, ranking y tabla) para
    los filtros dados. La usan tanto `/api/dashboard` (para pintar la
    pantalla) como `/api/exportar` (para que el Excel descargado muestre
    exactamente los mismos números y secciones que se ven en pantalla).
    Devuelve None cuando el filtro no deja ninguna campaña."""
    df_f = _filtrar(df, args)
    camp_f = _filtrar(df_camp, args)

    if camp_f.empty:
        return None

    total_destinatarios = numero(camp_f["destinatarios"].sum()) or 0
    total_evaluable = numero(camp_f["destinatarios_evaluables"].sum()) or 0
    total_efectivos = numero(camp_f["efectivos"].sum()) or 0
    efectividad_prom = round(total_efectivos / total_evaluable * 100, 2) if total_evaluable else None
    inversion_total = numero(camp_f["valor"].sum()) or 0
    costo_por_destinatario = round(inversion_total / total_destinatarios, 2) if total_destinatarios else None

    entidades_solicitantes = camp_f.explode("entidades", ignore_index=True)["entidades"].nunique()
    envios_por_tipo = df_f["tipo"].value_counts()
    ultimo_envio = df_f["fecha_evento"].max() if not df_f.empty else None

    kpis = {
        "total_campanas": entero(camp_f["campana_id"].nunique()),
        "total_envios": entero(df_f.shape[0]),
        "total_destinatarios": entero(total_destinatarios),
        "total_efectivos": entero(total_efectivos),
        "efectividad_promedio": efectividad_prom,
        "inversion_total": inversion_total,
        "costo_por_destinatario": costo_por_destinatario,
        "entidades_solicitantes": entero(entidades_solicitantes),
        "envios_sms": entero(envios_por_tipo.get("SMS", 0)),
        "envios_mailing": entero(envios_por_tipo.get("Mailing", 0)),
        "ultimo_envio": fecha_iso(ultimo_envio),
    }

    comparativo_canal = []
    for tipo, g in camp_f.groupby("tipo"):
        dst = numero(g["destinatarios"].sum()) or 0
        ev = numero(g["destinatarios_evaluables"].sum()) or 0
        efe = numero(g["efectivos"].sum()) or 0
        comparativo_canal.append({
            "tipo": tipo,
            "campanas": entero(g["campana_id"].nunique()),
            "destinatarios": entero(dst),
            "efectividad": round(efe / ev * 100, 2) if ev else None,
            "valor": numero(g["valor"].sum()),
        })

    evolucion = []
    for (orden, periodo), g in camp_f.groupby(["periodo_orden", "periodo"]):
        if orden == 0:
            continue
        dst = numero(g["destinatarios"].sum()) or 0
        ev = numero(g["destinatarios_evaluables"].sum()) or 0
        efe = numero(g["efectivos"].sum()) or 0
        evolucion.append({
            "periodo": periodo,
            "orden": int(orden),
            "campanas": entero(g["campana_id"].nunique()),
            "destinatarios": entero(dst),
            "efectividad": round(efe / ev * 100, 2) if ev else None,
        })
    evolucion.sort(key=lambda x: x["orden"])

    # Se "desempaca" por entidad individual (una campaña compartida, como
    # "Secretaría General / Prosperidad Social", cuenta para el ranking de
    # ambas) en vez de agrupar por el texto combinado de la campaña.
    explotado = camp_f.explode("entidades", ignore_index=True)
    top_entidades = []
    for entidad, g in explotado.groupby("entidades"):
        dst = numero(g["destinatarios"].sum()) or 0
        ev = numero(g["destinatarios_evaluables"].sum()) or 0
        efe = numero(g["efectivos"].sum()) or 0
        top_entidades.append({
            "entidad": entidad,
            "campanas": entero(g["campana_id"].nunique()),
            "destinatarios": entero(dst),
            "efectividad": round(efe / ev * 100, 2) if ev else None,
        })
    top_entidades.sort(key=lambda x: x["campanas"], reverse=True)
    top_entidades = top_entidades[:TOP_ENTIDADES]

    resultado_envios = {
        "efectivos": entero(camp_f["efectivos"].sum()),
        "no_entregados": entero(camp_f["no_entregados"].sum()),
        "rechazados": entero(camp_f["rechazados"].sum()),
        "expirados": entero(camp_f["expirados"].sum()),
        "invalidos": entero(camp_f["invalidos"].sum()),
    }

    con_efectividad = camp_f.dropna(subset=["efectividad"])
    mejores = con_efectividad.sort_values("efectividad", ascending=False).head(5)
    peores = con_efectividad.sort_values("efectividad", ascending=True).head(5)

    def a_ranking(fila):
        return {
            "campana": fila["campana"],
            "entidad": fila["entidad"],
            "tipo": fila["tipo"],
            "efectividad": numero(fila["efectividad"]),
            "destinatarios": entero(fila["destinatarios"]),
        }

    ranking = {
        "mejores": [a_ranking(f) for _, f in mejores.iterrows()],
        "peores": [a_ranking(f) for _, f in peores.iterrows()],
    }

    tabla_df = camp_f.sort_values("fecha_evento", ascending=False, na_position="last")
    columnas_tabla = ["campana", "entidad", "tipo", "mes", "destinatarios", "efectividad", "valor", "fecha_evento"]
    filas = []
    for _, f in tabla_df.iterrows():
        filas.append([
            f["campana"], f["entidad"], f["tipo"], f["mes"],
            entero(f["destinatarios"]), numero(f["efectividad"]), numero(f["valor"]),
            fecha_iso(f["fecha_evento"]),
        ])

    return {
        "seleccion": {
            "tipo": args.get("tipo", "Todos"),
            "entidad": _valores_seleccionados(args, "entidad", "Todas"),
            "mes": _valores_seleccionados(args, "mes", "Todos"),
            "anio": _valores_seleccionados(args, "anio", "Todos"),
            "buscar": args.get("buscar", ""),
        },
        "kpis": kpis,
        "comparativo_canal": comparativo_canal,
        "evolucion": evolucion,
        "evolucion_periodos": _evolucion_periodos(df_f),
        "top_entidades": top_entidades,
        "efectividad_entidades": _efectividad_por_entidad(df_f),
        "resultado_envios": resultado_envios,
        "ranking": ranking,
        "tabla": {
            "columnas_etiquetas": ["Campaña", "Entidad", "Canal", "Mes", "Destinatarios", "Efectividad %", "Valor", "Fecha"],
            "columnas": columnas_tabla,
            "filas": filas,
            "total_filas": len(filas),
        },
    }


def _filtros_disponibles(df_camp):
    anios = df_camp["anio_txt"].dropna().unique().tolist()
    anios.sort(key=lambda a: (a == "Sin año", a))
    return {
        "tipos": ["Todos"] + sorted(df_camp["tipo"].dropna().unique().tolist()),
        "entidades": ["Todas"] + _todas_las_entidades(df_camp),
        "meses": ["Todos"] + sorted(df_camp["mes"].dropna().unique().tolist(), key=lambda m: etl.MESES_ORDEN.index(m) if m in etl.MESES_ORDEN else 99),
        "anios": ["Todos"] + anios,
    }


# ------------------------------------------------------- modo genérico (respaldo) ---
def _clasificar_columnas(df):
    numericas, fechas, categoricas = [], [], []
    for col in df.columns:
        serie = df[col]
        if pd.api.types.is_datetime64_any_dtype(serie):
            fechas.append(col)
        elif pd.api.types.is_bool_dtype(serie):
            categoricas.append(col)
        elif pd.api.types.is_numeric_dtype(serie):
            numericas.append(col)
        else:
            categoricas.append(col)
    return {"numericas": numericas, "fechas": fechas, "categoricas": categoricas}


def _dashboard_generico(df, meta):
    args = request.args
    columnas = _clasificar_columnas(df)
    metrica = args.get("metrica") if args.get("metrica") in columnas["numericas"] else (columnas["numericas"][0] if columnas["numericas"] else None)
    agrupar = args.get("agrupar") if args.get("agrupar") in columnas["categoricas"] else (columnas["categoricas"][0] if columnas["categoricas"] else None)

    cifras = {"filas": int(len(df))}
    if metrica:
        cifras.update(total=numero(df[metrica].sum()), promedio=numero(df[metrica].mean()), maximo=numero(df[metrica].max()))

    por_categoria = []
    if agrupar:
        serie = df.groupby(agrupar)[metrica].sum() if metrica else df.groupby(agrupar).size()
        serie = serie.sort_values(ascending=False).head(MAX_CATEGORIAS_GENERICO)
        por_categoria = [{"etiqueta": str(k), "valor": numero(v)} for k, v in serie.items()]

    filas = [[celda(v) for v in fila] for fila in df.head(MAX_FILAS_TABLA_GENERICA).itertuples(index=False, name=None)]

    return jsonify(
        hay_datos=True, modo="generico", meta=meta,
        columnas=columnas,
        seleccion={"metrica": metrica, "agrupar": agrupar},
        cifras=cifras,
        por_categoria=por_categoria,
        tabla={"columnas": list(df.columns), "filas": filas, "total_filas": int(len(df))},
    )


@app.post("/api/subir")
@requiere_login
def subir():
    archivo = request.files.get("archivo")
    if archivo is None or not archivo.filename:
        return jsonify(error="Elige un archivo de Excel."), 400

    extension = Path(archivo.filename).suffix.lower()
    if extension not in EXTENSIONES:
        return jsonify(error="El archivo debe ser .xlsx o .xlsm."), 400

    temporal = ARCHIVO.parent / f".subida-{uuid.uuid4().hex}{extension}"
    try:
        archivo.save(temporal)
        with pd.ExcelFile(temporal) as xls:
            if not xls.sheet_names:
                raise ValueError("El archivo no tiene hojas.")
        os.replace(temporal, ARCHIVO)
        log.info("Archivo actualizado por subida: %s (%d bytes)", archivo.filename, ARCHIVO.stat().st_size)
    except Exception as exc:
        log.exception("Error al procesar archivo subido")
        return jsonify(error=f"No se pudo leer ese Excel: {exc}"), 422
    finally:
        temporal.unlink(missing_ok=True)

    return jsonify(ok=True, marca=marca_archivo())


def _descripcion_lista(etiqueta, valores, todas_txt):
    """'Entidad: Todas' si no hay nada seleccionado, o 'Entidad: A, B' con
    los valores elegidos (los filtros de entidad/mes/año admiten varios a
    la vez, como checkboxes)."""
    return f"{etiqueta}: {', '.join(valores) if valores else todas_txt}"


def _descripcion_filtros(seleccion):
    """Línea legible de los filtros aplicados, igual a la que se ve impresa
    en el encabezado del panel."""
    partes = [
        f"Canal: {seleccion['tipo']}",
        _descripcion_lista("Entidad", seleccion['entidad'], "Todas"),
        _descripcion_lista("Mes", seleccion['mes'], "Todos"),
        _descripcion_lista("Año", seleccion['anio'], "Todos"),
    ]
    if seleccion.get("buscar"):
        partes.append(f"Búsqueda: «{seleccion['buscar']}»")
    return " · ".join(partes)


def _imagen_desde_base64(data_url, ancho_max=760):
    """Decodifica una imagen PNG en base64 (como la devuelve
    `canvas.toBase64Image()` de Chart.js) a una imagen incrustable de
    openpyxl, escalada a un ancho razonable conservando la proporción.
    Devuelve None si `data_url` falta o no se puede decodificar."""
    if not data_url or not isinstance(data_url, str):
        return None
    try:
        from openpyxl.drawing.image import Image as ImagenXL

        _, _, b64 = data_url.partition(",")
        crudo = base64.b64decode(b64 or data_url)
        img = ImagenXL(BytesIO(crudo))
        if img.width and img.width > ancho_max:
            factor = ancho_max / img.width
            img.width = ancho_max
            img.height = int(img.height * factor)
        return img
    except Exception:
        log.exception("No se pudo incrustar una imagen de gráfico en el informe")
        return None


def _construir_libro_exportacion(datos, meta, imagenes=None):
    """Arma un libro de Excel con una hoja por cada sección del dashboard
    (resumen/KPIs, comparativo por canal, evolución mensual, entidades,
    resultado de envíos, ranking y la tabla detallada), con los MISMOS
    datos ya filtrados que se ven en pantalla — para que el informe
    descargado sea un reflejo fiel del panel, no solo la tabla cruda.

    `imagenes` (opcional): dict {"evolucion"|"canal"|"resultado"|"entidades":
    "data:image/png;base64,..."} con una captura exacta de cada gráfico tal
    como lo pintó Chart.js en el navegador. Cuando viene, esa imagen se
    incrusta tal cual (mismos colores y estilo que en pantalla) en vez de
    recrear el gráfico con el estilo nativo de Excel; si falta para algún
    gráfico, ese se arma con un gráfico nativo de respaldo."""
    imagenes = imagenes or {}
    from openpyxl import Workbook
    from openpyxl.chart import BarChart, LineChart, PieChart, Reference
    from openpyxl.formatting.rule import ColorScaleRule
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    # Rojo institucional de la Alcaldía de Bogotá (ver manual de marca) en
    # los encabezados de tabla, para que el informe exportado use la misma
    # paleta que el panel en pantalla.
    RELLENO_ENCABEZADO = PatternFill("solid", fgColor="E4032E")
    FUENTE_ENCABEZADO = Font(color="FFFFFF", bold=True)
    FUENTE_TITULO = Font(bold=True, size=13)
    FUENTE_ETIQUETA = Font(bold=True)
    RELLENO_FILA_PAR = PatternFill("solid", fgColor="FBEDED")
    BORDE_FINO = Border(
        left=Side(style="thin", color="E9D4D4"), right=Side(style="thin", color="E9D4D4"),
        top=Side(style="thin", color="E9D4D4"), bottom=Side(style="thin", color="E9D4D4"),
    )

    wb = Workbook()

    def nueva_hoja(nombre):
        hoja = wb.create_sheet(nombre)
        return hoja

    def fila_encabezado(hoja, fila, etiquetas):
        for col, texto in enumerate(etiquetas, start=1):
            celda = hoja.cell(row=fila, column=col, value=texto)
            celda.font = FUENTE_ENCABEZADO
            celda.fill = RELLENO_ENCABEZADO
            celda.alignment = Alignment(vertical="center")
        hoja.row_dimensions[fila].height = 20
        hoja.freeze_panes = hoja.cell(row=fila + 1, column=1).coordinate

    def ajustar_anchos(hoja, anchos):
        for col, ancho in enumerate(anchos, start=1):
            hoja.column_dimensions[get_column_letter(col)].width = ancho

    def aplicar_bordes(hoja, fila_inicio, fila_fin, col_inicio, col_fin):
        """Dibuja bordes finos alrededor de cada celda de una tabla, para que
        se vea como una tabla real y no como texto suelto."""
        for fila_n in range(fila_inicio, fila_fin + 1):
            for col_n in range(col_inicio, col_fin + 1):
                hoja.cell(row=fila_n, column=col_n).border = BORDE_FINO

    def aplicar_zebra(hoja, fila_inicio_datos, fila_fin, col_inicio, col_fin):
        """Sombrea de forma alterna las filas de datos (no el encabezado)
        para que tablas largas se lean con más facilidad."""
        for i, fila_n in enumerate(range(fila_inicio_datos, fila_fin + 1)):
            if i % 2 == 1:
                for col_n in range(col_inicio, col_fin + 1):
                    hoja.cell(row=fila_n, column=col_n).fill = RELLENO_FILA_PAR

    def aplicar_mapa_calor(hoja, col, fila_inicio, fila_fin):
        """Mapa de calor rojo→amarillo→verde sobre una columna de
        efectividad % (0 a 100), para distinguir de un vistazo las campañas
        o periodos flojos de los fuertes."""
        if fila_fin < fila_inicio:
            return
        letra = get_column_letter(col)
        rango = f"{letra}{fila_inicio}:{letra}{fila_fin}"
        regla = ColorScaleRule(
            start_type="num", start_value=0, start_color="F8696B",
            mid_type="num", mid_value=70, mid_color="FFEB84",
            end_type="num", end_value=100, end_color="63BE7B",
        )
        hoja.conditional_formatting.add(rango, regla)

    def tabla_con_estilo(hoja, fila_inicio, col_inicio, col_fin, fila_fin_datos, col_heatmap=None):
        """Aplica el acabado estándar (bordes + franjas alternas + mapa de
        calor opcional) a una tabla ya escrita, dado su rango de filas."""
        aplicar_bordes(hoja, fila_inicio, fila_fin_datos, col_inicio, col_fin)
        aplicar_zebra(hoja, fila_inicio + 1, fila_fin_datos, col_inicio, col_fin)
        if col_heatmap:
            aplicar_mapa_calor(hoja, col_heatmap, fila_inicio + 1, fila_fin_datos)

    # ---------------------------------------------------------- Resumen ---
    resumen = wb.active
    resumen.title = "Resumen"
    resumen["A1"] = "Línea 195 · Informe de campañas SMS & Mailing"
    resumen["A1"].font = FUENTE_TITULO
    resumen["A2"] = f"Archivo: {meta['nombre']}"
    resumen["A3"] = f"Generado: {datetime.now().astimezone().strftime('%Y-%m-%d %H:%M')}"
    resumen["A4"] = f"Filtros aplicados: {_descripcion_filtros(datos['seleccion'])}"
    for fila in (2, 3, 4):
        resumen.cell(row=fila, column=1).font = Font(italic=True, color="555555")

    etiquetas_kpi = {
        "total_campanas": "Total campañas",
        "total_envios": "Total envíos (olas)",
        "total_destinatarios": "Total destinatarios",
        "total_efectivos": "Total efectivos",
        "efectividad_promedio": "Efectividad promedio (%)",
        "inversion_total": "Inversión total ($)",
        "costo_por_destinatario": "Costo por destinatario ($)",
        "entidades_solicitantes": "Entidades solicitantes",
        "envios_sms": "Envíos por SMS",
        "envios_mailing": "Envíos por Mailing",
        "ultimo_envio": "Último envío medido",
    }
    fila_kpi_inicio = 6
    fila_encabezado(resumen, fila_kpi_inicio, ["Indicador", "Valor"])
    fila = fila_kpi_inicio
    for clave, etiqueta in etiquetas_kpi.items():
        fila += 1
        resumen.cell(row=fila, column=1, value=etiqueta)
        resumen.cell(row=fila, column=2, value=datos["kpis"].get(clave))
    tabla_con_estilo(resumen, fila_kpi_inicio, 1, 2, fila)

    fila += 2
    fila_resultado_titulo = fila
    fila_encabezado(resumen, fila_resultado_titulo, ["Resultado de envíos", "Cantidad"])
    etiquetas_resultado = {
        "efectivos": "Efectivos",
        "no_entregados": "No entregados",
        "rechazados": "Rechazados",
        "expirados": "Expirados",
        "invalidos": "Inválidos",
    }
    fila_resultado_inicio = fila + 1
    for clave, etiqueta in etiquetas_resultado.items():
        fila += 1
        resumen.cell(row=fila, column=1, value=etiqueta)
        resumen.cell(row=fila, column=2, value=datos["resultado_envios"].get(clave))
    fila_resultado_fin = fila
    tabla_con_estilo(resumen, fila_resultado_titulo, 1, 2, fila_resultado_fin)
    ajustar_anchos(resumen, [32, 46])

    # Gráfico "Resultado de envíos": se incrusta la captura exacta del
    # doughnut de pantalla cuando viene; si no, se arma uno nativo de torta.
    imagen_resultado = _imagen_desde_base64(imagenes.get("resultado"), ancho_max=480)
    if imagen_resultado is not None:
        resumen.add_image(imagen_resultado, "D6")
    else:
        torta_resultado = PieChart()
        torta_resultado.title = "Resultado de envíos"
        torta_resultado.height = 8
        torta_resultado.width = 12
        datos_torta = Reference(resumen, min_col=2, min_row=fila_resultado_inicio, max_row=fila_resultado_fin)
        categorias_torta = Reference(resumen, min_col=1, min_row=fila_resultado_inicio, max_row=fila_resultado_fin)
        torta_resultado.add_data(datos_torta, titles_from_data=False)
        torta_resultado.set_categories(categorias_torta)
        resumen.add_chart(torta_resultado, "D6")

    # ----------------------------------------------- Comparativo por canal ---
    canal = nueva_hoja("Comparativo por canal")
    fila_encabezado(canal, 1, ["Canal", "Campañas", "Destinatarios", "Efectividad %", "Valor"])
    for i, f in enumerate(datos["comparativo_canal"], start=2):
        canal.cell(row=i, column=1, value=f["tipo"])
        canal.cell(row=i, column=2, value=f["campanas"])
        canal.cell(row=i, column=3, value=f["destinatarios"])
        canal.cell(row=i, column=4, value=f["efectividad"])
        canal.cell(row=i, column=5, value=f["valor"])
    ajustar_anchos(canal, [14, 12, 16, 14, 16])
    tabla_con_estilo(canal, 1, 1, 5, max(1, len(datos["comparativo_canal"]) + 1), col_heatmap=4)

    imagen_canal = _imagen_desde_base64(imagenes.get("canal"))
    if imagen_canal is not None:
        canal.add_image(imagen_canal, "G2")
    else:
        ultima_fila_canal = len(datos["comparativo_canal"]) + 1
        if ultima_fila_canal > 1:
            grafico_canal = BarChart()
            grafico_canal.title = "Comparativo por canal"
            grafico_canal.y_axis.title = None
            grafico_canal.height = 8
            grafico_canal.width = 14
            # Igual que en pantalla: Campañas y Efectividad % en el mismo eje.
            datos_canal = Reference(canal, min_col=2, max_col=2, min_row=1, max_row=ultima_fila_canal)
            datos_canal_ef = Reference(canal, min_col=4, max_col=4, min_row=1, max_row=ultima_fila_canal)
            categorias_canal = Reference(canal, min_col=1, min_row=2, max_row=ultima_fila_canal)
            grafico_canal.add_data(datos_canal, titles_from_data=True)
            grafico_canal.add_data(datos_canal_ef, titles_from_data=True)
            grafico_canal.set_categories(categorias_canal)
            canal.add_chart(grafico_canal, "G2")

    # --------------------------------------------------- Evolución mensual ---
    evol = nueva_hoja("Evolución mensual")
    fila_encabezado(evol, 1, ["Periodo", "Campañas", "Destinatarios", "Efectividad %"])
    for i, f in enumerate(datos["evolucion"], start=2):
        evol.cell(row=i, column=1, value=f["periodo"])
        evol.cell(row=i, column=2, value=f["campanas"])
        evol.cell(row=i, column=3, value=f["destinatarios"])
        evol.cell(row=i, column=4, value=f["efectividad"])
    ajustar_anchos(evol, [20, 12, 16, 14])
    tabla_con_estilo(evol, 1, 1, 4, max(1, len(datos["evolucion"]) + 1), col_heatmap=4)

    imagen_evol = _imagen_desde_base64(imagenes.get("evolucion"))
    if imagen_evol is not None:
        evol.add_image(imagen_evol, "F2")
    else:
        ultima_fila_evol = len(datos["evolucion"]) + 1
        if ultima_fila_evol > 1:
            # Igual que en pantalla: barras de Campañas + línea de
            # Efectividad % en un eje secundario, para no aplastar la línea.
            barras_evol = BarChart()
            barras_evol.y_axis.axId = 100
            datos_campanas_evol = Reference(evol, min_col=2, max_col=2, min_row=1, max_row=ultima_fila_evol)
            barras_evol.add_data(datos_campanas_evol, titles_from_data=True)

            linea_evol = LineChart()
            linea_evol.y_axis.axId = 200
            linea_evol.y_axis.title = "Efectividad %"
            linea_evol.y_axis.crosses = "max"
            datos_efectividad_evol = Reference(evol, min_col=4, max_col=4, min_row=1, max_row=ultima_fila_evol)
            linea_evol.add_data(datos_efectividad_evol, titles_from_data=True)

            categorias_evol = Reference(evol, min_col=1, min_row=2, max_row=ultima_fila_evol)
            barras_evol.set_categories(categorias_evol)
            barras_evol.title = "Evolución mensual"
            barras_evol.height = 8
            barras_evol.width = 18
            barras_evol += linea_evol
            evol.add_chart(barras_evol, "F2")

    # ---------------------------------------------- Entidades (ranking) ---
    ents = nueva_hoja("Entidades")
    fila_encabezado(ents, 1, ["Entidad", "Campañas", "Destinatarios", "Efectividad %"])
    for i, f in enumerate(datos["top_entidades"], start=2):
        ents.cell(row=i, column=1, value=f["entidad"])
        ents.cell(row=i, column=2, value=f["campanas"])
        ents.cell(row=i, column=3, value=f["destinatarios"])
        ents.cell(row=i, column=4, value=f["efectividad"])
    ajustar_anchos(ents, [58, 12, 16, 14])
    tabla_con_estilo(ents, 1, 1, 4, max(1, len(datos["top_entidades"]) + 1), col_heatmap=4)

    imagen_ents = _imagen_desde_base64(imagenes.get("entidades"))
    if imagen_ents is not None:
        ents.add_image(imagen_ents, "F2")
    else:
        ultima_fila_ents = len(datos["top_entidades"]) + 1
        if ultima_fila_ents > 1:
            # Igual que en pantalla: barras horizontales de Campañas por entidad.
            grafico_ents = BarChart()
            grafico_ents.type = "bar"  # horizontal
            grafico_ents.title = "Entidades con más campañas"
            grafico_ents.height = 10
            grafico_ents.width = 18
            datos_ents = Reference(ents, min_col=2, max_col=2, min_row=1, max_row=ultima_fila_ents)
            categorias_ents = Reference(ents, min_col=1, min_row=2, max_row=ultima_fila_ents)
            grafico_ents.add_data(datos_ents, titles_from_data=True)
            grafico_ents.set_categories(categorias_ents)
            grafico_ents.legend = None
            ents.add_chart(grafico_ents, "F2")

    # --------------------------------------------------------- Ranking ---
    rank = nueva_hoja("Ranking")
    columnas_ranking = ["Campaña", "Entidad", "Canal", "Efectividad %", "Destinatarios"]

    def volcar_ranking(fila_inicio, titulo, filas_datos):
        rank.cell(row=fila_inicio, column=1, value=titulo).font = FUENTE_ETIQUETA
        fila_encabezado(rank, fila_inicio + 1, columnas_ranking)
        for i, f in enumerate(filas_datos, start=fila_inicio + 2):
            rank.cell(row=i, column=1, value=f["campana"])
            rank.cell(row=i, column=2, value=f["entidad"])
            rank.cell(row=i, column=3, value=f["tipo"])
            rank.cell(row=i, column=4, value=f["efectividad"])
            rank.cell(row=i, column=5, value=f["destinatarios"])
        fila_fin = fila_inicio + 1 + len(filas_datos)
        tabla_con_estilo(rank, fila_inicio + 1, 1, 5, fila_fin, col_heatmap=4)
        return fila_inicio + 2 + len(filas_datos)

    siguiente = volcar_ranking(1, "Mejores campañas (por efectividad)", datos["ranking"]["mejores"])
    volcar_ranking(siguiente + 1, "Peores campañas (por efectividad)", datos["ranking"]["peores"])
    ajustar_anchos(rank, [42, 36, 10, 14, 16])

    # ------------------------------------------------- Campañas (detalle) ---
    tabla = nueva_hoja("Campañas")
    fila_encabezado(tabla, 1, datos["tabla"]["columnas_etiquetas"])
    for i, f in enumerate(datos["tabla"]["filas"], start=2):
        for col, valor in enumerate(f, start=1):
            tabla.cell(row=i, column=col, value=valor)
    ajustar_anchos(tabla, [46, 44, 10, 12, 14, 14, 14, 12])
    num_cols_tabla = len(datos["tabla"]["columnas_etiquetas"])
    col_efectividad_tabla = datos["tabla"]["columnas_etiquetas"].index("Efectividad %") + 1
    tabla_con_estilo(
        tabla, 1, 1, num_cols_tabla, max(1, len(datos["tabla"]["filas"]) + 1),
        col_heatmap=col_efectividad_tabla,
    )

    for hoja in (canal, evol, ents, rank, tabla):
        hoja.sheet_view.showGridLines = False
    resumen.sheet_view.showGridLines = False

    return wb


def _fmt_num_es(valor, decimales=0):
    """Formatea un número al estilo es-CO: punto para miles, coma para
    decimales (p. ej. 2.848.965 o 94,50), igual que se ve en el panel."""
    if valor is None:
        return "—"
    try:
        n = float(valor)
    except (TypeError, ValueError):
        return "—"
    texto = f"{n:,.{decimales}f}"
    texto = texto.replace(",", "\x00").replace(".", ",").replace("\x00", ".")
    return texto


def _fmt_entero(valor):
    return _fmt_num_es(valor, 0)


def _fmt_moneda(valor):
    return "—" if valor is None else f"$ {_fmt_num_es(valor, 0)}"


def _fmt_pct(valor):
    return "—" if valor is None else f"{_fmt_num_es(valor, 2)} %"


def _fmt_fecha_larga(valor):
    if not valor:
        return "—"
    try:
        return datetime.strptime(valor, "%Y-%m-%d").strftime("%d/%m/%Y")
    except ValueError:
        return valor


def _color_mapa_calor_pdf(valor):
    """Mismo mapa de calor rojo→amarillo→verde (0/70/100) que el Excel
    exportado, pero como color de relleno de celda para reportlab."""
    from reportlab.lib import colors as rl_colors

    if valor is None:
        return rl_colors.HexColor("#F2F0F8")

    def hex_a_rgb(h):
        return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))

    def mezclar(c1, c2, t):
        return tuple(c1[i] + (c2[i] - c1[i]) * t for i in range(3))

    v = max(0.0, min(100.0, float(valor)))
    rojo, amarillo, verde = hex_a_rgb("F8696B"), hex_a_rgb("FFEB84"), hex_a_rgb("63BE7B")
    r, g, b = mezclar(rojo, amarillo, v / 70) if v <= 70 else mezclar(amarillo, verde, (v - 70) / 30)
    return rl_colors.Color(r / 255, g / 255, b / 255)


def _imagen_pdf_desde_base64(data_url, ancho_max_cm):
    """Decodifica una captura PNG en base64 (`canvas.toBase64Image()`) a un
    flowable Image de reportlab, escalada a un ancho máximo en cm
    conservando la proporción. Devuelve None si falta o no se puede leer."""
    if not data_url or not isinstance(data_url, str):
        return None
    try:
        from reportlab.lib.units import cm
        from reportlab.platypus import Image as ImagenPDF
        from PIL import Image as ImagenPIL

        _, _, b64 = data_url.partition(",")
        crudo = base64.b64decode(b64 or data_url)
        ancho_px, alto_px = ImagenPIL.open(BytesIO(crudo)).size
        ancho = ancho_max_cm * cm
        alto = ancho * (alto_px / ancho_px) if ancho_px else ancho
        return ImagenPDF(BytesIO(crudo), width=ancho, height=alto)
    except Exception:
        log.exception("No se pudo incrustar una imagen de gráfico en el PDF")
        return None


def _lectura_rapida_textos(datos):
    """Replica en el servidor las mismas 'conclusiones automáticas' que se
    ven en la tarjeta 'Lectura rápida' del panel (ver renderizarLecturaRapida
    en app.js), para que el PDF las incluya sin depender del navegador."""
    k = datos["kpis"]
    puntos = []
    mejor_canal = sorted(
        (c for c in datos["comparativo_canal"] if c["efectividad"] is not None),
        key=lambda c: c["efectividad"], reverse=True,
    )
    if mejor_canal:
        c = mejor_canal[0]
        puntos.append(f"<b>{c['tipo']}</b> tiene la mayor efectividad entre canales, con {_fmt_pct(c['efectividad'])}.")
    if datos["top_entidades"]:
        e = datos["top_entidades"][0]
        puntos.append(f"<b>{e['entidad']}</b> es la entidad con más campañas ({_fmt_entero(e['campanas'])}).")
    mejor_mes = sorted(
        (e for e in datos["evolucion"] if e["efectividad"] is not None),
        key=lambda e: e["efectividad"], reverse=True,
    )
    if mejor_mes:
        m = mejor_mes[0]
        puntos.append(f"<b>{m['periodo']}</b> fue el mes con mejor efectividad ({_fmt_pct(m['efectividad'])}).")
    if k["costo_por_destinatario"] is not None:
        puntos.append(f"Cada destinatario alcanzado cuesta en promedio <b>{_fmt_moneda(k['costo_por_destinatario'])}</b>.")
    return puntos or ["Sin datos suficientes todavía."]


def _construir_pdf_exportacion(datos, meta, imagenes=None):
    """Arma un PDF con las MISMAS secciones que se ven en el panel (KPIs,
    comparativo por canal, evolución, entidades, resultado de envíos,
    lectura rápida, ranking y la tabla detallada), igual que
    `_construir_libro_exportacion` arma el Excel.

    A diferencia de "Imprimir" del navegador (que depende de que Chrome
    repinte cada <canvas> a tiempo justo antes de imprimir, lo que a veces
    deja alguna gráfica en blanco), este PDF se genera siempre igual en el
    servidor con reportlab: mismos datos, mismas imágenes de gráfico ya
    capturadas por el navegador, sin ninguna carrera de tiempos. Las
    tarjetas de gráficas, KPI y ranking nunca se parten a la mitad entre
    hojas (se mueven completas a la siguiente página); la tabla "Campañas"
    sí se parte con normalidad entre hojas, con el encabezado repetido."""
    imagenes = imagenes or {}
    from reportlab.lib import colors as rl_colors
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
    from reportlab.lib.units import cm
    from reportlab.platypus import (Paragraph, SimpleDocTemplate, Spacer,
                                     Table, TableStyle)

    ROJO = rl_colors.HexColor("#E4032E")
    TEAL = rl_colors.HexColor("#1FAC8F")
    AMBAR = rl_colors.HexColor("#F9D248")
    GRAFITO = rl_colors.HexColor("#4A4A56")
    ROSA = rl_colors.HexColor("#E3272A")
    TEXTO = rl_colors.HexColor("#2B2540")
    SUAVE = rl_colors.HexColor("#736D8A")
    BORDE = rl_colors.HexColor("#E0DCEC")
    ZEBRA = rl_colors.HexColor("#FBEDED")
    BLANCO = rl_colors.white

    ANCHO_PAGINA, _ = A4
    MARGEN = 1.3 * cm
    ANCHO_UTIL_CM = (ANCHO_PAGINA - 2 * MARGEN) / cm
    # Relleno interno (izq+der) de una `tarjeta()`: al dimensionar una tabla
    # que va DENTRO de una tarjeta de ancho completo hay que restar esto,
    # o la tabla queda más ancha que el hueco disponible y se ve un
    # sobrante asomando por el borde derecho de la tarjeta.
    PAD_TARJETA_CM = 0.75

    estilos = getSampleStyleSheet()
    est_titulo = ParagraphStyle("TituloInforme", parent=estilos["Heading1"], fontSize=16, textColor=TEXTO, spaceAfter=2, fontName="Helvetica-Bold")
    est_subt = ParagraphStyle("Subtitulo", parent=estilos["Normal"], fontSize=8.5, textColor=SUAVE, spaceAfter=10)
    est_h2 = ParagraphStyle("H2", parent=estilos["Normal"], fontSize=11, textColor=TEXTO, fontName="Helvetica-Bold", spaceAfter=1)
    est_h2sub = ParagraphStyle("H2Sub", parent=estilos["Normal"], fontSize=8, textColor=SUAVE, spaceAfter=5)
    est_normal = ParagraphStyle("NormalInforme", parent=estilos["Normal"], fontSize=8.5, textColor=TEXTO, leading=11.5)
    est_kpi_etq = ParagraphStyle("KpiEtiqueta", parent=estilos["Normal"], fontSize=7.6, textColor=SUAVE, fontName="Helvetica-Bold")
    est_kpi_val = ParagraphStyle("KpiValor", parent=estilos["Normal"], fontSize=14, leading=17, textColor=TEXTO, fontName="Helvetica-Bold", spaceBefore=1, spaceAfter=2)
    est_kpi_nota = ParagraphStyle("KpiNota", parent=estilos["Normal"], fontSize=6.8, textColor=SUAVE, spaceBefore=2)
    est_th = ParagraphStyle("TH", parent=estilos["Normal"], fontSize=8, textColor=BLANCO, fontName="Helvetica-Bold")
    est_bullet = ParagraphStyle("Bullet", parent=estilos["Normal"], fontSize=8.5, textColor=TEXTO, leading=12, spaceAfter=6)

    buffer = BytesIO()
    doc = SimpleDocTemplate(
        buffer, pagesize=A4,
        leftMargin=MARGEN, rightMargin=MARGEN, topMargin=MARGEN, bottomMargin=MARGEN,
        title="Línea 195 · Panel de campañas",
    )
    story = []

    # ------------------------------------------------------- encabezado ---
    story.append(Paragraph("Línea 195 · Panel de campañas", est_titulo))
    generado = datetime.now().astimezone().strftime("%d/%m/%Y, %I:%M %p")
    filtros_txt = _descripcion_filtros(datos["seleccion"])
    story.append(Paragraph(f"Generado el {generado} · Fuente: {meta['nombre']} · {filtros_txt}", est_subt))
    regla = Table([[""]], colWidths=[ANCHO_UTIL_CM * cm], rowHeights=[0.03 * cm],
                  style=TableStyle([("BACKGROUND", (0, 0), (-1, -1), TEXTO)]))
    regla.hAlign = "LEFT"
    story.append(regla)
    story.append(Spacer(1, 10))

    def tarjeta(contenido, ancho_cm, alto_min_cm=None):
        """Envuelve una lista de flowables en una celda con borde fino, para
        que se vea como una tarjeta (igual que `.tarjeta` en el panel)."""
        estilo = [
            ("BOX", (0, 0), (-1, -1), 0.6, BORDE),
            ("LEFTPADDING", (0, 0), (-1, -1), 10), ("RIGHTPADDING", (0, 0), (-1, -1), 10),
            ("TOPPADDING", (0, 0), (-1, -1), 8), ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ]
        t = Table([[contenido]], colWidths=[ancho_cm * cm],
                   rowHeights=[alto_min_cm * cm] if alto_min_cm else None, style=TableStyle(estilo))
        t.hAlign = "LEFT"
        return t

    def fila_tarjetas(celdas_contenido, anchos_cm, espacio_cm=0.4):
        """Arma una fila de tarjetas lado a lado; si no caben completas en
        lo que queda de página, reportlab mueve la FILA entera (no la
        parte) a la siguiente, así ninguna tarjeta queda cortada a la
        mitad."""
        celdas = [tarjeta(c, a) for c, a in zip(celdas_contenido, anchos_cm)]
        t = Table([celdas], colWidths=[a * cm for a in anchos_cm],
                   style=TableStyle([
                       ("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (-1, -1), espacio_cm * cm),
                       ("TOPPADDING", (0, 0), (-1, -1), 0), ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
                       ("VALIGN", (0, 0), (-1, -1), "TOP"),
                   ]))
        t.hAlign = "LEFT"
        return t

    # -------------------------------------------------------------- KPIs ---
    k = datos["kpis"]
    kpi_items = [
        ("Campañas", _fmt_entero(k["total_campanas"]), f"{_fmt_entero(k['total_envios'])} envíos registrados"),
        ("Destinatarios alcanzados", _fmt_entero(k["total_destinatarios"]), f"{_fmt_entero(k['total_efectivos'])} mensajes efectivos"),
        ("Efectividad promedio", _fmt_pct(k["efectividad_promedio"]), "sobre envíos con resultado reportado"),
        ("Inversión total", _fmt_moneda(k["inversion_total"]), "suma de campañas filtradas"),
        ("Costo por destinatario", _fmt_moneda(k["costo_por_destinatario"]), "inversión ÷ destinatarios"),
    ]
    acentos = [ROJO, TEAL, AMBAR, GRAFITO, ROSA]
    ancho_kpi = (ANCHO_UTIL_CM - 0.4 * 2) / 3
    filas_kpi = [kpi_items[i:i + 3] for i in range(0, len(kpi_items), 3)]
    colores_filas = [acentos[i:i + 3] for i in range(0, len(acentos), 3)]
    for fila_items, fila_colores in zip(filas_kpi, colores_filas):
        celdas = []
        for (etq, val, nota), color in zip(fila_items, fila_colores):
            barra = Table([[""]], colWidths=[1.6 * cm], rowHeights=[0.12 * cm],
                           style=TableStyle([("BACKGROUND", (0, 0), (-1, -1), color)]))
            contenido = [barra, Spacer(1, 4), Paragraph(etq.upper(), est_kpi_etq),
                         Paragraph(val, est_kpi_val), Paragraph(nota, est_kpi_nota)]
            celdas.append(contenido)
        anchos = [ancho_kpi] * len(fila_items)
        story.append(fila_tarjetas(celdas, anchos))
        story.append(Spacer(1, 8))

    # --------------------------------------------------------- estadísticas ---
    stats = [
        ("Envíos por SMS", _fmt_entero(k["envios_sms"])),
        ("Envíos por mailing", _fmt_entero(k["envios_mailing"])),
        ("Entidades solicitantes", _fmt_entero(k["entidades_solicitantes"])),
        ("Último envío medido", _fmt_fecha_larga(k["ultimo_envio"])),
    ]
    ancho_stat = (ANCHO_UTIL_CM - 0.4 * 3) / 4
    celdas_stat = [[Paragraph(etq.upper(), est_kpi_etq), Paragraph(val, ParagraphStyle(
        "StatValor", parent=est_kpi_val, fontSize=11))] for etq, val in stats]
    story.append(fila_tarjetas(celdas_stat, [ancho_stat] * 4))
    story.append(Spacer(1, 12))

    # ------------------------------------------------------------ gráficos ---
    def tarjeta_grafico(titulo, subtitulo, clave_imagen, ancho_cm):
        img = _imagen_pdf_desde_base64(imagenes.get(clave_imagen), ancho_cm - 1.2)
        contenido = [Paragraph(titulo, est_h2), Paragraph(subtitulo, est_h2sub)]
        if img is not None:
            contenido.append(img)
        else:
            contenido.append(Paragraph("(Gráfico no disponible)", est_kpi_nota))
        return contenido

    ancho_2col = (ANCHO_UTIL_CM - 0.4) / 2
    story.append(fila_tarjetas([
        tarjeta_grafico("Evolución mensual", "Campañas y efectividad", "evolucion", ancho_2col),
        tarjeta_grafico("SMS vs. Mailing", "Volumen y efectividad", "canal", ancho_2col),
    ], [ancho_2col, ancho_2col]))
    story.append(Spacer(1, 8))

    story.append(fila_tarjetas([
        tarjeta_grafico("Resultado de envíos", "Distribución total", "resultado", ancho_2col),
        tarjeta_grafico("Entidades con más campañas", "Top por volumen", "entidades", ancho_2col),
    ], [ancho_2col, ancho_2col]))
    story.append(Spacer(1, 8))

    if imagenes.get("semanal"):
        story.append(fila_tarjetas([
            tarjeta_grafico("Mensajes efectivos por semana", "Evolución semanal por canal", "semanal", ANCHO_UTIL_CM),
        ], [ANCHO_UTIL_CM]))
        story.append(Spacer(1, 8))

    # -------------------------------------------------------- lectura rápida ---
    puntos = [Paragraph(f"✓ {p}", est_bullet) for p in _lectura_rapida_textos(datos)]
    story.append(tarjeta([Paragraph("Lectura rápida", est_h2), Spacer(1, 4)] + puntos, ANCHO_UTIL_CM))
    story.append(Spacer(1, 10))

    # -------------------------------------------- tabla con estilo genérica ---
    def tabla_con_bordes(encabezados, filas, anchos_cm, col_heatmap=None, repetir_encabezado=True):
        datos_tabla = [[Paragraph(h, est_th) for h in encabezados]] + filas
        estilo = [
            ("BACKGROUND", (0, 0), (-1, 0), ROJO),
            ("GRID", (0, 0), (-1, -1), 0.5, BORDE),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("LEFTPADDING", (0, 0), (-1, -1), 5), ("RIGHTPADDING", (0, 0), (-1, -1), 5),
            ("TOPPADDING", (0, 0), (-1, -1), 4), ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [BLANCO, ZEBRA]),
        ]
        t = Table(datos_tabla, colWidths=[a * cm for a in anchos_cm], repeatRows=1 if repetir_encabezado else 0)
        t.setStyle(TableStyle(estilo))
        t.hAlign = "LEFT"
        return t

    def celda_pct(valor):
        p = Paragraph(_fmt_pct(valor), ParagraphStyle("PctCelda", parent=est_normal, textColor=TEXTO))
        return p

    # ------------------------------------------------------ efectividad x entidad ---
    efect_ent = datos.get("efectividad_entidades") or []
    if efect_ent:
        filas = [[Paragraph(f"{e['entidad']} ({_fmt_entero(e['envios'])} envíos)", est_normal), celda_pct(e["efectividad"])]
                  for e in efect_ent]
        t = tabla_con_bordes(["Entidad", "Efectividad %"], filas, [ANCHO_UTIL_CM - PAD_TARJETA_CM - 3, 3])
        for i, e in enumerate(efect_ent, start=1):
            t.setStyle(TableStyle([("BACKGROUND", (1, i), (1, i), _color_mapa_calor_pdf(e["efectividad"]))]))
        story.append(tarjeta([Paragraph("Efectividad por entidad", est_h2), Spacer(1, 4), t], ANCHO_UTIL_CM))
        story.append(Spacer(1, 10))

    # -------------------------------------------------------------- rankings ---
    def tabla_ranking(filas_ranking):
        filas = []
        for i, f in enumerate(filas_ranking, start=1):
            filas.append([
                Paragraph(f"{i}. {f['campana'] or 'Sin nombre'}", est_normal),
                Paragraph(f"{f['entidad'] or '—'} · {f['tipo']}", ParagraphStyle("Meta", parent=est_kpi_nota, fontSize=7.2)),
                celda_pct(f["efectividad"]),
            ])
        t = Table(
            [[Paragraph("Campaña", est_th), Paragraph("", est_th), Paragraph("Efect. %", est_th)]] +
            [[f[0], f[1], f[2]] for f in filas],
            colWidths=[(ancho_2col - PAD_TARJETA_CM - 2.2) * cm, 0.01 * cm, 2.2 * cm], repeatRows=1,
        )
        estilo = [
            ("BACKGROUND", (0, 0), (-1, 0), ROJO), ("SPAN", (0, 0), (1, 0)),
            ("GRID", (0, 0), (-1, -1), 0.5, BORDE), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("LEFTPADDING", (0, 0), (-1, -1), 5), ("RIGHTPADDING", (0, 0), (-1, -1), 5),
            ("TOPPADDING", (0, 0), (-1, -1), 4), ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [BLANCO, ZEBRA]),
            ("SPAN", (0, 1), (1, 1)),
        ]
        for i in range(1, len(filas) + 1):
            estilo.append(("SPAN", (0, i), (1, i)))
        for i, f in enumerate(filas_ranking, start=1):
            estilo.append(("BACKGROUND", (2, i), (2, i), _color_mapa_calor_pdf(f["efectividad"])))
        t.setStyle(TableStyle(estilo))
        t.hAlign = "LEFT"
        return t

    if datos["ranking"]["mejores"] or datos["ranking"]["peores"]:
        story.append(fila_tarjetas([
            [Paragraph("Mejor desempeño", est_h2), Spacer(1, 4), tabla_ranking(datos["ranking"]["mejores"])],
            [Paragraph("Menor efectividad", est_h2), Spacer(1, 4), tabla_ranking(datos["ranking"]["peores"])],
        ], [ancho_2col, ancho_2col]))
        story.append(Spacer(1, 12))

    # -------------------------------------------------- Campañas (detalle) ---
    # A propósito esta tabla SÍ se deja partir con normalidad entre hojas
    # (a diferencia de las tarjetas de arriba): abarca demasiado texto como
    # para evitar el corte, y el encabezado se repite en cada página nueva.
    tabla_datos = datos["tabla"]
    story.append(Paragraph(f"Campañas — {tabla_datos['total_filas']} campaña(s), una fila por campaña", est_h2))
    story.append(Spacer(1, 4))
    etiquetas = tabla_datos["columnas_etiquetas"]
    col_efectividad = etiquetas.index("Efectividad %")
    anchos_detalle_cm = [3.6, 3.2, 1.6, 2.4, 2.4, 1.9, 2.1, 2.1]
    factor = ANCHO_UTIL_CM / sum(anchos_detalle_cm)
    anchos_detalle_cm = [a * factor for a in anchos_detalle_cm]

    filas_detalle = []
    valores_efectividad = []
    for f in tabla_datos["filas"]:
        valores_efectividad.append(f[col_efectividad])
        fila_fmt = []
        for idx, valor in enumerate(f):
            etiqueta = etiquetas[idx]
            if etiqueta == "Efectividad %":
                fila_fmt.append(Paragraph(_fmt_pct(valor), est_normal))
            elif etiqueta == "Valor":
                fila_fmt.append(Paragraph(_fmt_moneda(valor), est_normal))
            elif etiqueta == "Destinatarios":
                fila_fmt.append(Paragraph(_fmt_entero(valor), est_normal))
            elif etiqueta == "Fecha":
                fila_fmt.append(Paragraph(_fmt_fecha_larga(valor), est_normal))
            else:
                fila_fmt.append(Paragraph(str(valor) if valor is not None else "—", est_normal))
        filas_detalle.append(fila_fmt)

    tabla_grande = Table(
        [[Paragraph(h, est_th) for h in etiquetas]] + filas_detalle,
        colWidths=[a * cm for a in anchos_detalle_cm], repeatRows=1,
    )
    estilo_grande = [
        ("BACKGROUND", (0, 0), (-1, 0), ROJO),
        ("GRID", (0, 0), (-1, -1), 0.5, BORDE),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 4), ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ("TOPPADDING", (0, 0), (-1, -1), 3), ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [BLANCO, ZEBRA]),
    ]
    for i, valor in enumerate(valores_efectividad, start=1):
        estilo_grande.append(("BACKGROUND", (col_efectividad, i), (col_efectividad, i), _color_mapa_calor_pdf(valor)))
    tabla_grande.setStyle(TableStyle(estilo_grande))
    tabla_grande.hAlign = "LEFT"
    story.append(tabla_grande)

    doc.build(story)
    buffer.seek(0)
    return buffer


@app.route("/api/exportar", methods=["GET", "POST"])
@requiere_login
def exportar():
    """Exporta un Excel con las MISMAS secciones que se ven en el panel
    (resumen/KPIs, comparativo por canal, evolución, entidades, resultado
    de envíos, ranking y tabla detallada), aplicando los filtros vigentes.

    El frontend llama esto por POST con un cuerpo JSON
    `{"imagenes": {...}}`: una captura PNG exacta (`canvas.toBase64Image()`)
    de cada gráfico tal como lo pintó Chart.js, para incrustarla tal cual en
    el Excel en vez de recrear el gráfico con el estilo nativo de Excel. Por
    GET (sin cuerpo) se arman gráficos nativos de respaldo."""
    estado_cache = _cargar()
    if estado_cache["modo"] != "campanas":
        return jsonify(error="La exportación solo está disponible en modo campañas."), 400

    datos = _calcular_dashboard_campanas(estado_cache["df"], estado_cache["df_camp"], request.args)
    if datos is None:
        return jsonify(error="No hay campañas que coincidan con los filtros aplicados."), 400

    cuerpo = request.get_json(silent=True) or {}
    imagenes = cuerpo.get("imagenes") if isinstance(cuerpo.get("imagenes"), dict) else {}

    meta = {"nombre": ARCHIVO.name}
    wb = _construir_libro_exportacion(datos, meta, imagenes)
    ruta = BASE / "datos" / "_export_temp.xlsx"
    wb.save(ruta)
    return send_file(
        ruta,
        as_attachment=True,
        download_name="informe_campanas.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.route("/api/exportar-pdf", methods=["POST"])
@requiere_login
def exportar_pdf():
    """Genera un PDF de descarga directa con las MISMAS secciones del
    panel, construido siempre en el servidor con reportlab (ver
    `_construir_pdf_exportacion`). A diferencia de "Imprimir" del
    navegador -que depende de que Chrome alcance a congelar cada gráfico a
    tiempo antes de imprimir, y a veces deja alguno en blanco-, este PDF
    sale igual cada vez: el navegador solo aporta la captura PNG ya
    renderizada de cada gráfico (`canvas.toBase64Image()`), igual que para
    el Excel."""
    estado_cache = _cargar()
    if estado_cache["modo"] != "campanas":
        return jsonify(error="La exportación solo está disponible en modo campañas."), 400

    datos = _calcular_dashboard_campanas(estado_cache["df"], estado_cache["df_camp"], request.args)
    if datos is None:
        return jsonify(error="No hay campañas que coincidan con los filtros aplicados."), 400

    cuerpo = request.get_json(silent=True) or {}
    imagenes = cuerpo.get("imagenes") if isinstance(cuerpo.get("imagenes"), dict) else {}

    meta = {"nombre": ARCHIVO.name}
    try:
        buffer = _construir_pdf_exportacion(datos, meta, imagenes)
    except Exception:
        log.exception("No se pudo generar el PDF del panel")
        return jsonify(error="No se pudo generar el PDF."), 500

    return send_file(
        buffer,
        as_attachment=True,
        download_name="informe_campanas.pdf",
        mimetype="application/pdf",
    )


@app.errorhandler(413)
def archivo_grande(_):
    return jsonify(error="El archivo supera los 25 MB."), 413


@app.errorhandler(500)
def error_interno(exc):
    log.exception("Error interno no controlado")
    if request.path.startswith("/api/"):
        return jsonify(error="Ocurrió un error interno. Revisa los registros del servidor."), 500
    return "Ocurrió un error interno.", 500


if __name__ == "__main__":
    if not IRIS_PASSWORD:
        log.warning("IRIS_PASSWORD no está configurada: el panel queda SIN autenticación.")
    app.run(debug=DEBUG, host=os.environ.get("HOST", "127.0.0.1"), port=int(os.environ.get("PORT", 5000)))
