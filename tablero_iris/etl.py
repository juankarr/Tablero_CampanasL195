# -*- coding: utf-8 -*-
"""ETL especializado para la base de seguimiento de campañas SMS / Mailing (SGAM).

Toma el libro de Excel con las hojas "Mailing" y "SMS" y produce una tabla
unificada de campañas con columnas limpias y comparables entre los dos canales.
Tolera encabezados con espacios extra, fechas en español, columnas numéricas
guardadas como texto, filas de "olas" de envío con los datos descriptivos en
blanco (se heredan de la fila anterior) y nombres de entidad escritos de
formas distintas (tildes, mayúsculas, espacios).
"""
from __future__ import annotations

import logging
import re
import unicodedata
from datetime import datetime

import numpy as np
import pandas as pd

log = logging.getLogger(__name__)

MESES_ES = {
    "enero": 1, "febrero": 2, "marzo": 3, "abril": 4, "mayo": 5, "junio": 6,
    "julio": 7, "agosto": 8, "septiembre": 9, "setiembre": 9, "octubre": 10,
    "noviembre": 11, "diciembre": 12,
}
MESES_ORDEN = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio",
               "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]

RE_FECHA_ES = re.compile(
    r"(\d{1,2})\s+de\s+([a-záéíóúñ]+)\s+de\s+(\d{4})"
    r"(?:[^\d]{0,4}(\d{1,2}):(\d{2})\s*([ap])\.?\s*m\.?)?",
    re.IGNORECASE,
)


def _es_vacio(v):
    try:
        return bool(pd.isna(v))
    except (TypeError, ValueError):
        return False


def _norm_texto(s):
    if _es_vacio(s):
        return ""
    s = str(s).replace("\xa0", " ").replace(" ", " ").replace("​", "")
    s = re.sub(r"\s+", " ", s).strip()
    return s


def parse_fecha(valor):
    """Convierte fechas en español, texto libre o datetime nativo a Timestamp."""
    if _es_vacio(valor):
        return pd.NaT
    if isinstance(valor, (pd.Timestamp, datetime)):
        return pd.Timestamp(valor)
    s = _norm_texto(valor)
    if not s:
        return pd.NaT
    try:
        return pd.to_datetime(s, errors="raise")
    except Exception:
        pass
    m = RE_FECHA_ES.search(s.lower())
    if m:
        dia, mes_txt, anio, hora, minuto, ampm = m.groups()
        mes_num = MESES_ES.get(mes_txt)
        if mes_num:
            hora = int(hora) if hora else 0
            minuto = int(minuto) if minuto else 0
            if ampm:
                if ampm == "p" and hora != 12:
                    hora += 12
                if ampm == "a" and hora == 12:
                    hora = 0
            try:
                return pd.Timestamp(int(anio), mes_num, int(dia), hora, minuto)
            except Exception:
                return pd.NaT
    return pd.NaT


def parse_pct(valor):
    """Normaliza '% Efectividad' a un número 0-100, sin importar el formato original."""
    if _es_vacio(valor):
        return None
    if isinstance(valor, str):
        s = valor.strip().replace(",", ".")
        s = re.sub(r"%+", "", s).strip()
        if not s:
            return None
        try:
            v = float(s)
        except ValueError:
            return None
    else:
        try:
            v = float(valor)
        except (TypeError, ValueError):
            return None
    if v <= 1.5:
        v *= 100
    return round(v, 2)


def parse_valor(valor):
    """Normaliza la columna 'Valor' (costo) a float, tolerando texto sucio."""
    if _es_vacio(valor):
        return None
    if isinstance(valor, (int, float, np.integer, np.floating)):
        v = float(valor)
        return v if not pd.isna(v) else None
    s = re.sub(r"[^\d,.\-]", "", str(valor))
    s = s.replace(",", ".")
    if s.count(".") > 1:
        partes = s.split(".")
        s = "".join(partes[:-1]) + "." + partes[-1]
    try:
        return float(s) if s not in ("", "-", ".") else None
    except ValueError:
        return None


def _clave_entidad(nombre):
    s = _norm_texto(nombre)
    s = s.lstrip("|/ ").rstrip(". ")
    s_sin_tildes = "".join(
        c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn"
    )
    return s_sin_tildes.lower().strip()


_PALABRAS_VACIAS = {"de", "la", "el", "los", "las", "y", "del", "a", "distrital", "secretaria"}

# Dependencias/áreas que reportan como entidades distintas pero en realidad
# pertenecen a la Secretaría General de la Alcaldía Mayor de Bogotá: cualquier
# nombre de entidad que contenga alguno de estos fragmentos (ya normalizado,
# sin tildes y en minúsculas) se canoniza directamente a "Secretaría General",
# sin pasar por el agrupamiento por frecuencia ni por la fusión difusa.
_DEPENDENCIAS_SECRETARIA_GENERAL = [
    "secretaria general",
    "consejeria distrital de paz",
    "subsecretaria",
]

# Área/dependencia específica dentro de la Secretaría General, para que no se
# pierda al unificar: el nombre final muestra tanto la entidad ("Secretaría
# General Alcaldía Mayor de Bogotá") como el área ("... - Subsecretaría").
# Variantes de redacción distintas de la misma área (p. ej. con o sin el
# prefijo "Dirección de Reparación Integral") se tratan como una sola.
_AREAS_SECRETARIA_GENERAL = [
    ("consejeria distrital de paz", "Oficina Consejería Distrital de Paz, Víctimas y Reconciliación"),
    ("subsecretaria", "Subsecretaría"),
    ("observatorio distrital de victimas", "Observatorio Distrital de Víctimas del Conflicto Armado"),
]


def _canon_entidad_individual(clave):
    """Si `clave` (ya normalizada vía _clave_entidad) corresponde a la
    Secretaría General o a alguna de sus dependencias conocidas, devuelve el
    nombre canónico fijo. En caso contrario devuelve None."""
    for fragmento in _DEPENDENCIAS_SECRETARIA_GENERAL:
        if fragmento in clave:
            return "Secretaría General"
    return None


def _area_secretaria_general(clave):
    """Si `clave` pertenece a la familia de la Secretaría General Y además
    menciona una dependencia/área específica conocida, devuelve el nombre de
    esa área (p. ej. "Subsecretaría"). Si es la Secretaría General "pura"
    (sin área identificable) devuelve None, para no agregar un sufijo
    vacío."""
    if not any(fragmento in clave for fragmento in _DEPENDENCIAS_SECRETARIA_GENERAL):
        return None
    for fragmento, area in _AREAS_SECRETARIA_GENERAL:
        if fragmento in clave:
            return area
    return None


def _tokens_significativos(clave):
    return {t for t in clave.split() if t not in _PALABRAS_VACIAS and len(t) > 2}


# Catálogo oficial de entidades: lista maestra de nombres "como deben ser".
# Cuando una entidad detectada en el Excel coincide (por nombre completo, por
# sigla, o por contención de palabras clave) con una de estas, el nombre final
# que ve el usuario es siempre el de este catálogo, sin importar cómo haya
# llegado escrita en el archivo origen.
_CATALOGO_ENTIDADES = [
    "Administradora Colombiana de Pensiones - Colpensiones",
    "Agencia Distrital para la Educación Superior, La Ciencia y la Tecnología - Atenea",
    "Área Limpia D.C. S.A.S. E.S.P.",
    "Bogotá Limpia S.A.S E.S.P.",
    "Caja de Vivienda Popular - CVP",
    "Canal Capital",
    "Capital Salud EPS-S",
    "Ciudad Limpia Bogotá S.A ESP",
    "Comisión Nacional del Servicio Civil - CNSC",
    "Contraloría de Bogotá",
    "Departamento Administrativo de la Defensoría del Espacio Público - DADEP",
    "Departamento Administrativo del Servicio Civil Distrital - DASCD",
    "Camára de Comercio de Bogotá - CCB",
    "Departamento Administrativo para la Prosperidad Social.",
    "Dirección de Impuestos y Aduanas Nacionales - DIAN",
    "Empresa de Acueducto y Alcantarillado de Bogotá ESP",
    "Empresa de Renovación y Desarrollo Urbano de Bogotá - Renobo",
    "Empresa de Telecomunicaciones de Bogotá S.A. - ETB - ESP",
    "Empresa de Transporte del Tercer Milenio Transmilenio S.A.",
    "Empresa Metro de Bogotá S.A.",
    "ENEL COLOMBIA",
    "Entidad de Gestión Administrativa y Técnica -EGAT",
    "Fondo de Prestaciones Económicas, Cesantías y Pensiones - FONCEP",
    "Fondo Nacional del Ahorro - FNA",
    "Fundación Gilberto Alzate Avendaño - FUGA",
    "Instituto Colombiano de Bienestar Familiar - ICBF",
    "Instituto de Desarrollo Urbano - IDU",
    "Instituto Distrital de Ciencia, Biotecnología Innovación en Salud - IDCBIS",
    "Instituto Distrital de Gestión de Riesgos y Cambio Climático - IDIGER",
    "Instituto Distrital de la Participación y Acción Comunal - IDPAC",
    "Instituto Distrital de las Artes - IDARTES",
    "Instituto Distrital de Patrimonio Cultural - IDPC",
    "Instituto Distrital de Protección y Bienestar Animal - IDPYBA",
    "Instituto Distrital de Recreación y Deporte - IDRD",
    "Instituto Distrital de Turismo - IDT",
    "Instituto Distrital para la Protección de la Niñez y la Juventud - IDIPRON",
    "Instituto Para la Economía Social - IPES",
    "Instituto para la Investigación Educativa y Desarrollo Pedagógico - IDEP",
    "Jardín Botánico de Bogotá José Celestino Mutis",
    "Limpieza Metropolitana S.A. ESP",
    "Ministerio de la Igualdad",
    "Lotería de Bogotá",
    "Ministerio de Relaciones Exteriores",
    "Orquesta Filarmónica de Bogotá - OFB",
    "Personería de Bogotá",
    "Positiva Compañía de Seguros",
    "Promoambiental Distrito S.A.S. ESP",
    "Registraduría Distrital del Estado Civil",
    "Secretaría de Educación del Distrito - SED",
    "Secretaría Distrital de Ambiente - SDA",
    "Secretaría Distrital de Cultura, Recreación y Deporte - SDCRD",
    "Secretaría Distrital de Desarrollo Económico - SDDE",
    "Secretaría Distrital de Gobierno - SDG",
    "Secretaría Distrital de Hacienda - SDH",
    "Secretaría Distrital de Integración Social - SDIS",
    "Secretaría Distrital de la Mujer",
    "Secretaría Distrital de Movilidad - SDM",
    "Secretaría Distrital de Planeación - SDP",
    "Secretaría Distrital de Salud - SDS",
    "Secretaría Distrital de Seguridad, Convivencia y Justicia - SDSCJ",
    "Secretaría Distrital del Hábitat",
    "Secretaría General Alcaldía Mayor de Bogotá",
    "Secretaría Jurídica Distrital - SJD",
    "Servicio Nacional de Aprendizaje - SENA",
    "Soporte Portal Bogotá",
    "Subred Integrada de Servicios de Salud Centro Oriente E.S.E.",
    "Subred Integrada de Servicios de Salud Norte E.S.E.",
    "Subred Integrada de Servicios de Salud Sur E.S.E.",
    "Subred Integrada de Servicios de Salud Sur Occidente E.S.E.",
    "Terminal de Transportes S.A.",
    "Unidad Administrativa Especial Cuerpo Oficial de Bomberos - UAECOB",
    "Unidad Administrativa Especial de Catastro Distrital - UAECD",
    "Unidad Administrativa Especial de Gestión Pensional y Parafiscales de la Protección Social - UGPP",
    "Unidad Administrativa Especial de Rehabilitación y Mantenimiento Vial - UAERMV",
    "Unidad Administrativa Especial de Servicios Públicos - UAESP",
    "Unidad Administrativa Especial Migración Colombia",
    "Universidad Colegio Mayor de Cundinamarca",
    "Universidad Distrital Francisco Jose de Caldas",
    "Universidad Militar Nueva Granada",
    "Vanti S.A. ESP",
    "Veeduría Distrital",
    "REVAL S.A.S.",
]


def _preparar_catalogo_entidades():
    """Pre-procesa el catálogo oficial una sola vez: para cada entidad separa
    el nombre descriptivo de su sigla (lo que sigue al último " - "), y
    calcula sus claves/tokens normalizados para poder comparar contra los
    nombres que llegan del Excel."""
    catalogo = []
    for oficial in _CATALOGO_ENTIDADES:
        # Se toma el SEGUNDO segmento (no el último) como sigla: una entrada
        # como "Empresa de Telecomunicaciones de Bogotá S.A. - ETB - ESP"
        # tiene dos separadores " - ", y la sigla real es "ETB", no el
        # sufijo genérico "ESP" (Empresa de Servicios Públicos) que comparten
        # muchas otras entidades de la lista.
        partes = oficial.split(" - ")
        if len(partes) >= 2:
            nombre_largo, sigla = partes[0], partes[1]
        else:
            nombre_largo, sigla = oficial, None
        clave_largo = _clave_entidad(nombre_largo)
        catalogo.append({
            "oficial": oficial,
            "clave_largo": clave_largo,
            "tokens": _tokens_significativos(clave_largo),
            "clave_sigla": _clave_entidad(sigla) if sigla else None,
        })
    return catalogo


_CATALOGO = _preparar_catalogo_entidades()


def _contiene_como_unidad(clave, sub):
    """True si `sub` aparece en `clave` como unidad propia (no incrustado
    dentro de otra palabra): por ejemplo "sdis" en "secretaria sdis" sí, pero
    no una coincidencia casual dentro de una palabra más larga. Tolera que
    `sub` esté pegado a un guion bajo u otro separador no alfabético, como en
    "...pensiones_de_bogota_foncep"."""
    patron = re.compile(r"(?<![a-zñ])" + re.escape(sub) + r"(?![a-zñ])")
    return patron.search(clave) is not None


def _nombre_oficial_entidad(nombre):
    """Busca en el catálogo oficial la entidad que mejor corresponde a
    `nombre` (por sigla, por nombre completo exacto, por contención del uno
    en el otro, o por superposición de palabras clave) y devuelve su nombre
    "como debe ser". Si no hay una coincidencia razonable, devuelve None (y
    quien llama conserva el nombre ya normalizado tal cual)."""
    clave = _clave_entidad(nombre)
    if not clave:
        return None
    tokens = _tokens_significativos(clave)

    mejor_oficial = None
    mejor_score = 0
    for entrada in _CATALOGO:
        score = 0
        if entrada["clave_sigla"] and _contiene_como_unidad(clave, entrada["clave_sigla"]):
            score = 1000
        elif clave == entrada["clave_largo"]:
            score = 900
        elif entrada["clave_largo"] and len(clave) >= 6 and (
            clave in entrada["clave_largo"] or entrada["clave_largo"] in clave
        ):
            score = 500 + min(len(clave), len(entrada["clave_largo"]))
        elif len(tokens) >= 2 and len(entrada["tokens"]) >= 2 and (
            tokens <= entrada["tokens"] or entrada["tokens"] <= tokens
        ):
            score = 100 + len(tokens & entrada["tokens"])
        if score > mejor_score:
            mejor_score = score
            mejor_oficial = entrada["oficial"]
    return mejor_oficial


def _fusionar_variantes_cercanas(nombres_por_frecuencia):
    """Segunda pasada: une entidades cuyo conjunto de palabras clave de una
    está contenido en el de otra (p. ej. 'Secretaría de Desarrollo Económico'
    y 'Secretaría Distrital de Desarrollo Económico').

    Exige al menos 2 palabras clave en AMBOS lados antes de fusionar: con un
    solo token en común (p. ej. 'jurídica') se fusionarían por error entidades
    sin relación, como 'Secretaría Jurídica Distrital' dentro de 'Dirección
    Distrital de Política Jurídica'."""
    grupos = list(nombres_por_frecuencia)  # ya viene ordenado por frecuencia desc
    mapa_final = {}
    representantes = []  # (clave_tokens, nombre_canonico)
    for nombre in grupos:
        tokens = _tokens_significativos(_clave_entidad(nombre))
        fusion = None
        if len(tokens) >= 2:
            for tk_rep, nombre_rep in representantes:
                if len(tk_rep) >= 2 and (tokens <= tk_rep or tk_rep <= tokens):
                    fusion = nombre_rep
                    break
        if fusion:
            mapa_final[nombre] = fusion
        else:
            mapa_final[nombre] = nombre
            representantes.append((tokens, nombre))
    return mapa_final


def normalizar_entidades(serie):
    """Agrupa variantes de escritura de la misma entidad (tildes, mayúsculas,
    espacios, puntuación, y nombres casi idénticos) y devuelve el nombre
    canónico más frecuente de cada grupo. Las dependencias conocidas de la
    Secretaría General (ver `_DEPENDENCIAS_SECRETARIA_GENERAL`) se canonizan
    directamente, sin importar su frecuencia relativa. Como último paso,
    cualquier nombre que coincida con el catálogo oficial de entidades (ver
    `_CATALOGO_ENTIDADES`) se reemplaza por el nombre "como debe ser" de ese
    catálogo; lo que no coincida con ninguna entidad conocida conserva el
    nombre ya normalizado, sin forzarlo a encajar en el catálogo."""
    limpio = serie.map(_norm_texto).map(lambda s: s.lstrip("|/ ").rstrip(". "))
    claves = limpio.map(_clave_entidad)
    vacios = claves == ""

    overrides = claves.map(_canon_entidad_individual)
    aplica_override = overrides.notna() & ~vacios
    sin_override = (~aplica_override) & (~vacios)

    if sin_override.any():
        conteo_exacto = limpio[sin_override].groupby(claves[sin_override]).agg(lambda x: x.value_counts().idxmax())
        resultado = claves.map(conteo_exacto)
    else:
        resultado = pd.Series([None] * len(serie), index=serie.index, dtype=object)
    resultado = resultado.mask(aplica_override, overrides)

    frecuencias = resultado[~vacios].value_counts()
    mapa_fusion = _fusionar_variantes_cercanas(frecuencias.index.tolist())
    resultado = resultado.map(lambda n: mapa_fusion.get(n, n))

    # Último paso: todo nombre que el catálogo oficial reconozca se muestra
    # siempre con el nombre exacto del catálogo, sin importar cuál de sus
    # variantes "ganó" en los pasos anteriores.
    nombres_unicos = resultado[~vacios].unique().tolist()
    mapa_oficial = {n: (_nombre_oficial_entidad(n) or n) for n in nombres_unicos}
    resultado = resultado.map(lambda n: mapa_oficial.get(n, n))

    # Si la fila original mencionaba un área/dependencia específica de la
    # Secretaría General (p. ej. "Subsecretaría"), se agrega al nombre final
    # en vez de perderse al unificar todo bajo la entidad: "Secretaría
    # General Alcaldía Mayor de Bogotá - Subsecretaría".
    areas = claves.map(_area_secretaria_general)
    resultado = pd.Series(
        [f"{nombre} - {area}" if (pd.notna(area) and pd.notna(nombre)) else nombre
         for nombre, area in zip(resultado, areas)],
        index=resultado.index,
    )

    resultado[vacios] = None
    return resultado


def normalizar_entidades_multiples(serie):
    """Igual que `normalizar_entidades`, pero soporta celdas que mencionan
    varias entidades a la vez, separadas por "/" (p. ej. 'Secretaría General
    /Prosperidad Social'). Devuelve una tupla `(texto, listas)`:

    - `texto`: Series con el nombre a mostrar (entidades canónicas únicas,
      unidas con " / " cuando hay más de una).
    - `listas`: Series donde cada elemento es la lista (sin duplicados) de
      entidades canónicas mencionadas en esa fila, útil para filtrar y para
      que una campaña con dos entidades cuente en el ranking de ambas."""
    partes_por_fila = serie.map(_norm_texto).map(
        lambda s: [p.strip() for p in s.split("/") if p.strip()] if s else []
    )

    todas = [p for partes in partes_por_fila for p in partes]
    if not todas:
        texto_vacio = pd.Series([None] * len(serie), index=serie.index)
        listas_vacias = pd.Series([[] for _ in range(len(serie))], index=serie.index)
        return texto_vacio, listas_vacias

    canon_todas = normalizar_entidades(pd.Series(todas))
    mapa = dict(zip(todas, canon_todas.tolist()))

    def _canonizar_lista(partes):
        vistas = []
        for p in partes:
            c = mapa.get(p)
            if c and c not in vistas:
                vistas.append(c)
        return vistas

    listas = partes_por_fila.map(_canonizar_lista)
    texto = listas.map(lambda l: " / ".join(l) if l else None)
    return texto, listas


def _col(df, *nombres_posibles):
    mapa = {c.strip().lower(): c for c in df.columns}
    for n in nombres_posibles:
        real = mapa.get(n.strip().lower())
        if real is not None:
            return df[real]
    return pd.Series([None] * len(df), index=df.index)


def _calcular_efectividad(efectivos, total_base):
    return np.where(
        (total_base.fillna(0) > 0) & efectivos.notna(),
        (efectivos / total_base * 100).round(2),
        np.nan,
    )


def _combinar_efectividad(efectividad_raw, efectivos, total_base):
    """Usa el porcentaje de efectividad reportado en el archivo cuando es un
    valor plausible (0-100); si falta, o es claramente erróneo —por ejemplo,
    en el archivo nuevo una fila trae en esa celda el conteo de efectivos en
    vez de un porcentaje (850 en lugar de ~85%)—, se recalcula a partir de
    destinatarios y efectivos reales en lugar de propagar el dato dañado."""
    calc = pd.Series(_calcular_efectividad(efectivos, total_base), index=efectivos.index)
    raw_usable = efectividad_raw.where(efectividad_raw.notna() & (efectividad_raw <= 100))
    return raw_usable.where(raw_usable.notna(), calc)


def _ffill_texto(serie, es_continuacion):
    """Rellena hacia abajo los campos descriptivos SOLO en las filas marcadas
    como 'ola adicional' de un mismo envío (Cosecutivo en blanco). Una fila
    con Cosecutivo nuevo nunca hereda texto de la campaña anterior, aunque
    su propio campo venga vacío."""
    limpio = serie.map(_norm_texto)
    para_ffill = limpio.mask(es_continuacion)
    ancla = para_ffill.ffill()
    return limpio.where(~es_continuacion, ancla).fillna("")


def _preparar_hoja(df, tipo):
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    df = df.dropna(how="all")
    if df.empty:
        return pd.DataFrame()

    cosecutivo_bruto = pd.to_numeric(_col(df, "cosecutivo"), errors="coerce")
    es_continuacion = cosecutivo_bruto.isna()
    es_continuacion.iloc[0] = False  # la primera fila nunca es continuación
    cosecutivo = cosecutivo_bruto.ffill()
    campana_id = tipo[:1] + "-" + cosecutivo.fillna(-1).astype(int).astype(str)

    campana_txt = _ffill_texto(_col(df, "nombre campana"), es_continuacion)
    entidad_bruta = _ffill_texto(_col(df, "entidad solicitante"), es_continuacion)
    entidad_txt, entidades_lista = normalizar_entidades_multiples(entidad_bruta)
    caracterizacion_txt = _ffill_texto(_col(df, "caracterización"), es_continuacion)
    objetivo_txt = _ffill_texto(_col(df, "objetivo"), es_continuacion)
    mes_txt = _ffill_texto(_col(df, "mes"), es_continuacion).str.title()

    fecha_solicitud = _col(df, "fecha solicitud").map(parse_fecha)
    fecha_realizacion = _col(df, "fecha realización de campaña").map(parse_fecha)
    if tipo == "Mailing":
        fecha_envio = _col(df, "fecha de envio").map(parse_fecha)
        hora_envio = _col(df, "hora de envio").map(_norm_texto)
    else:
        fecha_envio = _col(df, "fecha   y hora del envio  masivo", "fecha y hora del envio masivo").map(parse_fecha)
        hora_envio = pd.Series([""] * len(df), index=df.index)

    total_base = pd.to_numeric(_col(df, "total base de datos", "registros entregados  en la base original", "registros entregados en la base original"), errors="coerce")
    validos = pd.to_numeric(_col(df, "registros validos para envio", "registros validos", "registros válidos"), errors="coerce")
    invalidos = pd.to_numeric(_col(df, "registros invalidos", "registros errados"), errors="coerce")
    efectivos = pd.to_numeric(_col(df, "mensajes efectivos", "efectivos"), errors="coerce")
    expirados = pd.to_numeric(_col(df, "expirados"), errors="coerce")
    rechazados = pd.to_numeric(_col(df, "rechazados"), errors="coerce")
    no_entregados = pd.to_numeric(_col(df, "no entregados", "mensajes no efectivos"), errors="coerce")
    total_mensajes = pd.to_numeric(_col(df, "suma de mensajes", "total general"), errors="coerce")

    valor = _col(df, "valor").map(parse_valor)
    efectividad = _col(df, "% efectividad").map(parse_pct)
    efectividad_final = _combinar_efectividad(efectividad, efectivos, total_base)

    # "Suma de Mensajes"/"Total general" son segmentos facturados (un SMS largo
    # puede partirse en varios), NO destinatarios: se guardan aparte y NUNCA se
    # usan como denominador de efectividad ni como volumen de audiencia.
    segmentos_facturados = total_mensajes

    # Volumen real de audiencia = base de datos objetivo de cada ola de envío.
    # Se toma el máximo contra la suma de resultados individuales porque en
    # algunas filas la base quedó mal digitada (menor que sus propios resultados).
    suma_resultados = efectivos.fillna(0) + no_entregados.fillna(0) + expirados.fillna(0) + rechazados.fillna(0) + invalidos.fillna(0)
    destinatarios = np.where(
        total_base.notna() | (suma_resultados > 0),
        np.fmax(total_base.fillna(0), suma_resultados),
        np.nan,
    )
    destinatarios = pd.Series(destinatarios, index=df.index)
    # Para la EFECTIVIDAD solo cuentan las olas donde sabemos cuántos mensajes
    # fueron efectivos; una ola sin ese dato no debe diluir el promedio.
    destinatarios_evaluables = destinatarios.where(efectivos.notna())

    out = pd.DataFrame({
        "campana_id": campana_id,
        "tipo": tipo,
        "entidad": entidad_txt,
        "entidades": entidades_lista,
        "campana": campana_txt,
        "caracterizacion": caracterizacion_txt,
        "objetivo": objetivo_txt,
        "mes": mes_txt,
        "fecha_solicitud": fecha_solicitud,
        "fecha_realizacion": fecha_realizacion,
        "fecha_envio": fecha_envio,
        "hora_envio": hora_envio,
        "total_base": total_base,
        "validos": validos,
        "invalidos": invalidos,
        "efectivos": efectivos,
        "expirados": expirados,
        "rechazados": rechazados,
        "no_entregados": no_entregados,
        "destinatarios": destinatarios,
        "destinatarios_evaluables": destinatarios_evaluables,
        "segmentos_facturados": segmentos_facturados,
        "valor": valor,
        "efectividad": efectividad_final,
        "informe_enviado": _col(df, "informe enviado").map(_norm_texto).str.upper(),
    })

    return _derivar_periodo(out)


def _derivar_periodo(out):
    """Calcula fecha_evento (mejor fecha disponible), mes/año/periodo y su
    orden cronológico. Comparte esta lógica tanto el parser del formato
    histórico (hojas Mailing/SMS separadas) como el del formato unificado
    nuevo (una sola hoja SMS con columna "Servicio"), para que ambos
    produzcan columnas de periodo consistentes entre sí."""
    fecha_evento = out["fecha_envio"]
    fecha_evento = fecha_evento.where(fecha_evento.notna(), out["fecha_realizacion"])
    fecha_evento = fecha_evento.where(fecha_evento.notna(), out["fecha_solicitud"])
    out["fecha_evento"] = fecha_evento

    mes_num = out["fecha_evento"].dt.month
    mes_desde_fecha = mes_num.map(lambda m: MESES_ORDEN[int(m) - 1] if pd.notna(m) else None)
    mes_vacio = out["mes"].isna() | (out["mes"] == "") | (out["mes"] == "Nan")
    out.loc[mes_vacio, "mes"] = mes_desde_fecha[mes_vacio]
    out["mes"] = out["mes"].fillna("Sin mes")

    anio = out["fecha_evento"].dt.year
    out["mes_num"] = mes_num.fillna(0).astype(int)
    out["anio"] = anio
    out["anio_txt"] = anio.map(lambda v: str(int(v)) if pd.notna(v) else "Sin año")
    out["periodo_orden"] = np.where(anio.notna(), anio.fillna(0).astype(int) * 100 + out["mes_num"], 0)
    out["periodo"] = np.where(anio.notna(), out["mes"] + " " + anio.astype("Int64").astype(str), out["mes"])

    return out


def _preparar_hoja_unificada(df):
    """Parser para el formato nuevo de la base: una sola hoja "SMS" que
    mezcla envíos de SMS y de Mailing, distinguidos por la columna
    "Servicio". Mapea los encabezados nuevos (incluidos un par con errores
    de tipeo en el archivo original: "Contenodo del Mensaje",
    "Contactactos efectivos") a las mismas columnas canónicas que produce
    `_preparar_hoja` para el formato histórico."""
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    df = df.dropna(how="all")
    if df.empty:
        return pd.DataFrame()

    servicio_txt = _col(df, "servicio").map(_norm_texto).str.upper()
    tipo = pd.Series(
        np.where(servicio_txt.str.contains("MAILING", na=False), "Mailing", "SMS"),
        index=df.index,
    )

    consecutivo = pd.to_numeric(_col(df, "consecutivo"), errors="coerce")
    campana_id = tipo.str[:1] + "-" + consecutivo.fillna(-1).astype(int).astype(str)

    campana_txt = _col(df, "nombre de la campaña", "nombre de la campana").map(_norm_texto)
    entidad_bruta = _col(df, "entidad").map(_norm_texto)
    entidad_txt, entidades_lista = normalizar_entidades_multiples(entidad_bruta)
    # El contenido del mensaje es lo más parecido a un "objetivo"/texto de
    # búsqueda en este formato; no hay columna equivalente a "caracterización".
    objetivo_txt = _col(df, "contenodo del mensaje", "contenido del mensaje").map(_norm_texto)
    caracterizacion_txt = pd.Series([""] * len(df), index=df.index)

    fecha_solicitud = _col(df, "fecha de la solicitud").map(parse_fecha)
    fecha_prueba = _col(df, "fecha de la prueba").map(parse_fecha)
    # "Fecha envio masivo" es la fecha real de envío; cuando falta, se usa la
    # fecha de la prueba (preparación del envío masivo) como mejor estimado.
    fecha_envio = _col(df, "fecha envio masivo", "fecha envió masivo", "fecha de envio masivo", "fecha de envió masivo").map(parse_fecha)
    fecha_envio = fecha_envio.where(fecha_envio.notna(), fecha_prueba)
    fecha_realizacion = _col(df, "fecha de la campaña", "fecha de la campana").map(parse_fecha)
    # Nota: la columna "Fecha de envio" (última del archivo) es, en la
    # práctica, una fecha administrativa de cierre de informe posterior al
    # evento (129/212 filas quedan después de "Fecha de la campaña"), así que
    # deliberadamente no se usa para calcular fecha_evento.

    total_base = pd.to_numeric(_col(df, "# registros base original"), errors="coerce")
    validos = pd.to_numeric(_col(df, "# registros base depurada"), errors="coerce")
    efectivos = pd.to_numeric(_col(df, "contactactos efectivos", "contactos efectivos"), errors="coerce")
    expirados = pd.to_numeric(_col(df, "expirados"), errors="coerce")
    rechazados = pd.to_numeric(_col(df, "rechazados"), errors="coerce")
    no_entregados = pd.to_numeric(_col(df, "no entregados"), errors="coerce")
    # Igual que en el formato histórico: "Suma de Mensajes Efectivos" son
    # segmentos facturados (un SMS largo se puede partir en varios), NO el
    # número real de destinatarios. Difiere de "Contactactos efectivos" en
    # más de la mitad de las filas, así que deben mantenerse separados.
    segmentos_facturados = pd.to_numeric(_col(df, "suma de mensajes efectivos"), errors="coerce")

    # No hay columna de "registros inválidos" en este formato; se infiere
    # como la diferencia entre la base original y la base depurada.
    invalidos = (total_base - validos).clip(lower=0)
    invalidos = invalidos.where(total_base.notna() & validos.notna())

    valor = _col(df, "valor").map(parse_valor)
    efectividad = _col(df, "% efectividad").map(parse_pct)
    efectividad_final = _combinar_efectividad(efectividad, efectivos, total_base)

    suma_resultados = efectivos.fillna(0) + no_entregados.fillna(0) + expirados.fillna(0) + rechazados.fillna(0) + invalidos.fillna(0)
    destinatarios = np.where(
        total_base.notna() | (suma_resultados > 0),
        np.fmax(total_base.fillna(0), suma_resultados),
        np.nan,
    )
    destinatarios = pd.Series(destinatarios, index=df.index)
    destinatarios_evaluables = destinatarios.where(efectivos.notna())

    out = pd.DataFrame({
        "campana_id": campana_id,
        "tipo": tipo,
        "entidad": entidad_txt,
        "entidades": entidades_lista,
        "campana": campana_txt,
        "caracterizacion": caracterizacion_txt,
        "objetivo": objetivo_txt,
        # Las dos columnas crudas "Mes" del archivo nuevo son ambiguas (el
        # mapeo de columnas por nombre normalizado no puede distinguir "Mes "
        # de "Mes", y en 32/212 filas tienen valores distintos), así que el
        # mes se deriva siempre de fecha_evento en _derivar_periodo.
        "mes": pd.Series([""] * len(df), index=df.index),
        "fecha_solicitud": fecha_solicitud,
        "fecha_realizacion": fecha_realizacion,
        "fecha_envio": fecha_envio,
        "hora_envio": pd.Series([""] * len(df), index=df.index),
        "total_base": total_base,
        "validos": validos,
        "invalidos": invalidos,
        "efectivos": efectivos,
        "expirados": expirados,
        "rechazados": rechazados,
        "no_entregados": no_entregados,
        "destinatarios": destinatarios,
        "destinatarios_evaluables": destinatarios_evaluables,
        "segmentos_facturados": segmentos_facturados,
        "valor": valor,
        "efectividad": efectividad_final,
        "informe_enviado": pd.Series([""] * len(df), index=df.index),
    })

    return _derivar_periodo(out)


def cargar_campanas(ruta_excel):
    """Lee el libro completo y devuelve (df_unificado, hojas_usadas).

    Soporta dos formatos de origen:
    - Histórico: hojas separadas "Mailing" y "SMS".
    - Nuevo/unificado: una sola hoja "SMS" que mezcla ambos canales,
      distinguidos por una columna "Servicio" (se detecta buscando esa
      columna en la hoja "SMS"; si está presente se usa este formato aunque
      también exista una hoja "Mailing" residual)."""
    with pd.ExcelFile(ruta_excel) as xls:
        hojas = xls.sheet_names
        partes = []
        usadas = []

        sms_df = None
        if "SMS" in hojas:
            sms_df = pd.read_excel(xls, sheet_name="SMS")
            sms_df.columns = [str(c).strip() for c in sms_df.columns]

        columnas_sms = {c.strip().lower() for c in sms_df.columns} if sms_df is not None else set()
        formato_unificado = "servicio" in columnas_sms

        if formato_unificado:
            if sms_df is not None and not sms_df.dropna(how="all").empty:
                partes.append(_preparar_hoja_unificada(sms_df))
                usadas.append("SMS (formato unificado)")
        else:
            if "Mailing" in hojas:
                m = pd.read_excel(xls, sheet_name="Mailing")
                if not m.dropna(how="all").empty:
                    partes.append(_preparar_hoja(m, "Mailing"))
                    usadas.append("Mailing")
            if sms_df is not None and not sms_df.dropna(how="all").empty:
                partes.append(_preparar_hoja(sms_df, "SMS"))
                usadas.append("SMS")

    if not partes:
        return None, []

    df = pd.concat(partes, ignore_index=True)
    df = df[df["campana"].astype(str).str.strip() != ""]
    log.info("ETL: %d filas de envío, %d campañas únicas, hojas=%s", len(df), df["campana_id"].nunique(), usadas)
    return df.reset_index(drop=True), usadas


def es_formato_campanas(ruta_excel):
    """Detecta si el Excel tiene la estructura esperada (hojas Mailing/SMS)."""
    try:
        with pd.ExcelFile(ruta_excel) as xls:
            return "Mailing" in xls.sheet_names or "SMS" in xls.sheet_names
    except Exception:
        return False
