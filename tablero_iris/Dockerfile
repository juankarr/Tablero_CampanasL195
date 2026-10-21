# Línea 195 · Panel de campañas SMS & Mailing
# Imagen lista para producción: sirve la app con gunicorn (no el servidor
# de desarrollo de Flask) y corre como usuario sin privilegios.

FROM python:3.11-slim

# Evita archivos .pyc y fuerza salida sin buffer (logs visibles al instante).
ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1

WORKDIR /app

# Instala primero solo las dependencias para aprovechar el cache de capas
# de Docker: si el código cambia pero no requirements.txt, esta capa no
# se vuelve a construir.
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copia el resto del proyecto.
COPY . .

# Usuario sin privilegios: nunca correr la app como root.
RUN useradd --create-home --uid 1000 iris \
    && mkdir -p /app/datos \
    && chown -R iris:iris /app
USER iris

EXPOSE 8000

# Comprobación de salud: /salud no requiere sesión, así que funciona igual
# con o sin IRIS_PASSWORD configurada.
HEALTHCHECK --interval=30s --timeout=5s --start-period=10s --retries=3 \
    CMD python -c "import urllib.request,sys; sys.exit(0) if urllib.request.urlopen('http://127.0.0.1:8000/salud', timeout=3).status==200 else sys.exit(1)"

# gunicorn: servidor WSGI de producción. 2 workers es un punto de partida
# razonable para un panel de uso interno; ajusta con la variable WEB_CONCURRENCY
# si hace falta más capacidad.
CMD ["gunicorn", "--bind", "0.0.0.0:8000", "--workers", "2", "--timeout", "60", "app:app"]
