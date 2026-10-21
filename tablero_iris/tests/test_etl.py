# -*- coding: utf-8 -*-
"""Pruebas del ETL contra los problemas reales encontrados en la base de
seguimiento de campañas: fechas en español, porcentajes en formatos
distintos, costos con texto sucio, entidades escritas de formas diferentes,
y filas de "ola adicional" que no deben mezclar campañas distintas."""
import sys
from pathlib import Path

import pandas as pd
import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import etl


class TestParseFecha:
    def test_fecha_nativa_datetime(self):
        import datetime
        assert etl.parse_fecha(datetime.datetime(2025, 9, 12)) == pd.Timestamp("2025-09-12")

    def test_fecha_iso_texto(self):
        assert etl.parse_fecha("2025-09-12") == pd.Timestamp("2025-09-12")

    def test_fecha_espanol_con_hora_pm(self):
        r = etl.parse_fecha("15 de septiembre de 2025 4:39 p. m.")
        assert r == pd.Timestamp("2025-09-15 16:39")

    def test_fecha_espanol_con_hora_am(self):
        r = etl.parse_fecha("16 de septiembre de 2025 11:14 a. m.")
        assert r == pd.Timestamp("2025-09-16 11:14")

    def test_fecha_espanol_narrow_nbsp(self):
        # espacio Unicode "narrow no-break space" que aparece en algunas filas
        r = etl.parse_fecha("23 de septiembre de 2025 11:44 a. m.")
        assert r == pd.Timestamp("2025-09-23 11:44")

    def test_fecha_espanol_sin_hora(self):
        r = etl.parse_fecha("30 de septiembre de 2025")
        assert r == pd.Timestamp("2025-09-30")

    def test_fecha_vacia(self):
        assert pd.isna(etl.parse_fecha(None))
        assert pd.isna(etl.parse_fecha(float("nan")))
        assert pd.isna(etl.parse_fecha(""))

    def test_fecha_basura(self):
        assert pd.isna(etl.parse_fecha("no es una fecha"))

    def test_mediodia_12pm_no_se_duplica(self):
        r = etl.parse_fecha("1 de enero de 2025 12:00 p. m.")
        assert r == pd.Timestamp("2025-01-01 12:00")

    def test_medianoche_12am(self):
        r = etl.parse_fecha("1 de enero de 2025 12:00 a. m.")
        assert r == pd.Timestamp("2025-01-01 00:00")


class TestParsePct:
    def test_texto_con_simbolo_porcentaje(self):
        assert etl.parse_pct("85.33 %") == 85.33

    def test_fraccion_decimal(self):
        assert etl.parse_pct(0.969) == 96.9

    def test_doble_simbolo_porcentaje_sucio(self):
        assert etl.parse_pct("97%%") == 97.0

    def test_coma_decimal(self):
        assert etl.parse_pct("85,33%") == 85.33

    def test_ya_en_escala_100(self):
        assert etl.parse_pct(85.33) == 85.33

    def test_vacio(self):
        assert etl.parse_pct(None) is None
        assert etl.parse_pct(float("nan")) is None
        assert etl.parse_pct("") is None

    def test_texto_no_numerico(self):
        assert etl.parse_pct("n/a") is None


class TestParseValor:
    def test_numero_directo(self):
        assert etl.parse_valor(29826.86) == 29826.86

    def test_entero(self):
        assert etl.parse_valor(5000) == 5000.0

    def test_texto_con_simbolos(self):
        assert etl.parse_valor("$29.826,86") is not None  # no debe reventar

    def test_vacio(self):
        assert etl.parse_valor(None) is None
        assert etl.parse_valor(float("nan")) is None

    def test_guion_como_vacio(self):
        assert etl.parse_valor("-") is None


class TestNormalizarEntidades:
    def test_agrupa_tildes_y_mayusculas(self):
        serie = pd.Series([
            "Secretaria Distrital de Integración Social",
            "Secretaría Distrital de Integración social",
            "SECRETARIA DISTRITAL DE INTEGRACION SOCIAL",
        ])
        resultado = etl.normalizar_entidades(serie)
        assert resultado.nunique() == 1

    def test_fusiona_variantes_con_y_sin_distrital(self):
        serie = pd.Series([
            "Secretaría de Desarrollo Económico",
            "Secretaría de Desarrollo Económico",
            "Secretaría Distrital de Desarrollo Económico",
        ])
        resultado = etl.normalizar_entidades(serie)
        assert resultado.nunique() == 1

    def test_entidades_distintas_no_se_fusionan(self):
        serie = pd.Series(["Secretaría General", "FONCEP", "Prosperidad Social"])
        resultado = etl.normalizar_entidades(serie)
        assert resultado.nunique() == 3

    def test_vacios_quedan_como_none(self):
        serie = pd.Series(["Secretaría General", "", None])
        resultado = etl.normalizar_entidades(serie)
        assert resultado.iloc[1] is None or pd.isna(resultado.iloc[1])
        assert resultado.iloc[2] is None or pd.isna(resultado.iloc[2])


class TestFfillTexto:
    def test_hereda_solo_en_continuacion(self):
        serie = pd.Series(["Campaña A", "", "Campaña B", ""])
        es_continuacion = pd.Series([False, True, False, True])
        resultado = etl._ffill_texto(serie, es_continuacion)
        assert list(resultado) == ["Campaña A", "Campaña A", "Campaña B", "Campaña B"]

    def test_fila_nueva_vacia_no_hereda(self):
        # Cosecutivo propio (no es continuación) pero el campo vino vacío:
        # no debe robar el nombre de la campaña anterior.
        serie = pd.Series(["Campaña A", "Campaña A", ""])
        es_continuacion = pd.Series([False, True, False])
        resultado = etl._ffill_texto(serie, es_continuacion)
        assert resultado.iloc[2] == ""


class TestIntegracionArchivoReal:
    """Estas pruebas corren contra el Excel real si está disponible; si no,
    se saltan (para que la suite funcione en cualquier entorno)."""

    RUTA = Path(__file__).resolve().parent.parent / "datos" / "datos.xlsx"

    @classmethod
    @pytest.fixture(scope="class")
    def df(cls):
        if not cls.RUTA.exists():
            pytest.skip("no hay datos.xlsx de referencia en este entorno")
        df, hojas = etl.cargar_campanas(cls.RUTA)
        assert df is not None
        return df

    def test_detecta_formato_campanas(self):
        if not self.RUTA.exists():
            pytest.skip("no hay datos.xlsx de referencia en este entorno")
        assert etl.es_formato_campanas(self.RUTA) is True

    def test_no_hay_campanas_vacias(self, df):
        assert (df["campana"].astype(str).str.strip() == "").sum() == 0

    def test_efectivos_nunca_superan_destinatarios(self, df):
        con_datos = df.dropna(subset=["efectivos", "destinatarios"])
        imposibles = con_datos[con_datos["efectivos"] > con_datos["destinatarios"]]
        assert len(imposibles) == 0

    def test_efectividad_en_rango_valido(self, df):
        con_efectividad = df.dropna(subset=["efectividad"])
        assert (con_efectividad["efectividad"] >= 0).all()
        assert (con_efectividad["efectividad"] <= 100.01).all()

    def test_entidades_normalizadas_reducen_variantes(self, df):
        # antes de normalizar había 33 variantes de escritura para ~14 entidades reales
        assert df["entidad"].nunique() < 20

    def test_cada_campana_tiene_un_solo_nombre_o_variantes_de_la_misma(self, df):
        """En el formato histórico, un mismo campana_id (Consecutivo) siempre
        tenía un único nombre de campaña. En el formato nuevo, cada "ola" de
        un mismo Consecutivo puede traer su propio nombre con el detalle de
        la sub-audiencia (p. ej. "...Ciclo 7 _Jóvenes" vs
        "...Ciclo 7 _Persona Mayor"), ya que ahora ese campo nunca llega en
        blanco. Lo que de verdad debe protegerse es que no se mezclen por
        error dos campañas sin relación bajo el mismo id: exigimos que,
        cuando hay más de un nombre, todos compartan un prefijo común
        razonablemente largo."""
        import os

        for campana_id, nombres in df.groupby("campana_id")["campana"]:
            unicos = nombres.unique().tolist()
            if len(unicos) <= 1:
                continue
            minusculas = [n.lower() for n in unicos]
            prefijo_comun = os.path.commonprefix(minusculas)
            largo_minimo = min(8, len(min(minusculas, key=len)) - 1)
            assert len(prefijo_comun) >= largo_minimo, (
                f"{campana_id}: nombres de campaña sin relación aparente: {unicos}"
            )
