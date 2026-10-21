# Línea 195 · Panel de campañas SMS & Mailing

Dashboard profesional en Flask para analizar las campañas de envío de SMS y
correo (Mailing) de una base de seguimiento en Excel: cuántas campañas se
hicieron, a cuántas personas llegaron realmente, qué tan efectivas fueron y
cuánto costaron — con filtros por canal, entidad, mes y año.

## Características

- **ETL robusto**: limpia fechas en texto libre en español, porcentajes en
  formatos mixtos, valores monetarios, y unifica nombres de entidades que
  llegan escritos de formas distintas (mayúsculas, tildes, abreviaturas,
  dependencias/áreas de una misma entidad). Reconoce automáticamente tanto el
  formato histórico (hojas separadas "Mailing" y "SMS") como el formato
  unificado más reciente (una sola hoja "SMS" con columna "Servicio" que
  distingue el canal). Distingue destinatarios reales de segmentos
  facturados, para no inflar ni desinflar la efectividad reportada, y separa
  por entidad las campañas que se solicitaron conjuntamente entre dos
  entidades (p. ej. "Secretaría General / Prosperidad Social") para que cada
  una aparezca correctamente en sus propios filtros y rankings. Cada fila de
  la base de Excel se trata como una campaña independiente (no se agrupan
  ni se suman filas que compartan "Consecutivo"), y la fecha de cada evento
  se toma de la columna "Fecha envió masivo" (o de la fecha de la prueba si
  esa falta).
- **Panel visual** con KPIs, evolución mensual, comparativo por canal,
  mensajes efectivos por semana/mes (con selector de métrica: mensajes
  efectivos, envíos o valor), ranking de entidades por volumen y por
  efectividad, y una tabla detallada con orden por columna y paginación.
- **Filtros**: canal (una opción a la vez); entidad, mes y año con checkboxes
  de selección múltiple (se puede combinar, por ejemplo, dos o más meses, o
  varias entidades, a la vez); y búsqueda de texto libre.
- **Identidad visual institucional**: paleta de colores de la Alcaldía de
  Bogotá (según el manual de uso de marca ciudad Bogotá), con el rojo del
  escudo como color principal (barra lateral, botones, acentos) sobre
  tarjetas blancas, apoyado por el amarillo institucional; el azul casi no
  se usa. Incluye la franja roja bajo el encabezado que usan los portales
  oficiales de la Alcaldía. Aplicada tanto en el panel como en el informe
  Excel exportado.
- **Modo oscuro** con persistencia en el navegador y sin parpadeo al cargar.
- **Autenticación** por usuario y contraseña compartidos (opcional: si se
  deja `IRIS_PASSWORD` vacía, el panel no pide login).
- **Exportar**: botón de descarga directa de un PDF con las mismas secciones
  del panel (KPIs, comparativo por canal, evolución mensual, entidades,
  resultado de envíos, lectura rápida, ranking y la tabla detallada de
  campañas), con los filtros vigentes. El PDF se arma siempre en el
  servidor (con reportlab) a partir de los mismos datos que se ven en
  pantalla y de una captura exacta de cada gráfico (`canvas.toBase64Image`)
  — no depende de que el navegador "congele" los gráficos a tiempo antes de
  imprimir, así que sale igual de completo y con los mismos colores cada
  vez, a diferencia de usar Ctrl+P / Imprimir del navegador directamente.
  Las tarjetas de gráficos, KPI y ranking nunca quedan cortadas a la mitad
  entre hojas (saltan completas a la página siguiente); la tabla
  "Campañas" sí se parte con normalidad entre hojas, con el encabezado
  repetido. También exporta un Excel (.xlsx) con una hoja por cada sección
  del panel, con los mismos filtros y números que se ven en pantalla. Los
  gráficos se incrustan como capturas con fondo claro y texto oscuro
  (legibles sobre el blanco de la hoja, en vez del degradado de color que
  se ve en pantalla), y solo si no se pueden capturar el Excel recurre a
  gráficos nativos equivalentes como respaldo. Todas las tablas de ambos
  informes llevan bordes finos, franjas alternas para facilitar la
  lectura, y un mapa de calor (rojo → amarillo → verde) sobre cada columna
  de "Efectividad %" para ubicar de un vistazo las campañas o periodos más
  fuertes y más flojos.
- **Pruebas automatizadas** del ETL (34 pruebas con pytest) que protegen los
  casos difíciles de los datos reales.
- **Listo para Docker**, con variables de entorno y registro (logging)
  estructurado.

## Requisitos

- Python 3.11+ (probado con 3.11)
- Un Excel con hojas "Mailing" y "SMS" (formato histórico) o una sola hoja
  "SMS" con columna "Servicio" que distinga SMS/Mailing (formato nuevo
  unificado) — o cualquier otro Excel: el panel cae a un modo genérico de
  exploración de datos si no reconoce ninguno de los dos formatos

## Instalación local

```bash
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
cp .env.example .env    # y edita IRIS_USUARIO / IRIS_PASSWORD / SECRET_KEY
python3 app.py
```

Abre `http://127.0.0.1:5000`. Por defecto el acceso es usuario `SGAM195` y
contraseña `Linea195` — cámbialos en `.env` antes de exponer el panel. Si
dejaste `IRIS_PASSWORD` vacío en `.env`, el panel no pide login (útil solo
para pruebas locales).

### Variables de entorno

Ver `.env.example` para la lista completa con comentarios. Las más
importantes:

| Variable        | Para qué sirve                                            |
|-----------------|------------------------------------------------------------|
| `IRIS_USUARIO`  | Usuario de acceso. Por defecto `SGAM195`.                   |
| `IRIS_PASSWORD` | Contraseña de acceso. Por defecto `Linea195`. Vacía = sin login. |
| `SECRET_KEY`    | Firma la cookie de sesión. Genera una fija en producción.   |
| `ARCHIVO_EXCEL` | Ruta del Excel fuente (por defecto `datos/datos.xlsx`).     |
| `LOG_LEVEL`     | `DEBUG`, `INFO`, `WARNING` o `ERROR`.                       |
| `FLASK_DEBUG`   | `1` solo en desarrollo; nunca en un servidor expuesto.      |

## Actualizar los datos

Desde el panel: botón **"Actualizar datos"** en la barra lateral, sube un
`.xlsx`/`.xlsm` y se reemplaza el archivo activo sin reiniciar el servidor.

## Pruebas

```bash
pip install -r requirements-dev.txt
pytest tests/ -v
```

Las pruebas cubren tanto funciones aisladas del ETL (parseo de fechas,
porcentajes, valores, normalización de entidades) como invariantes contra el
archivo Excel real (por ejemplo: la efectividad nunca supera 100%, ningún
`campana_id` mezcla dos campañas distintas).

## Despliegue con Docker

```bash
cp .env.example .env    # define IRIS_PASSWORD y SECRET_KEY
docker compose up --build
```

Esto sirve el panel en `http://localhost:8000` con `gunicorn` (no el
servidor de desarrollo de Flask). El Excel vive en `./datos`, montado como
volumen, así que puedes reemplazarlo sin reconstruir la imagen.

Sin `docker compose`, equivalente manual:

```bash
docker build -t tablero-iris .
docker run -p 8000:8000 \
  -e IRIS_PASSWORD=tu-clave \
  -e SECRET_KEY=$(python3 -c "import secrets; print(secrets.token_hex(32))") \
  -v $(pwd)/datos:/app/datos \
  tablero-iris
```

El contenedor expone `/salud` (sin autenticación) para chequeos de salud de
tu orquestador.

## Estructura del proyecto

```
app.py              Backend Flask: rutas, autenticación, agregaciones
etl.py               Limpieza y transformación de los datos del Excel
templates/           Vistas Jinja2 (panel y login)
static/              CSS, JS y librerías vendorizadas (Chart.js, Font Awesome)
tests/               Pruebas automatizadas del ETL
datos/               Excel activo del panel
Dockerfile, docker-compose.yml, .dockerignore   Despliegue en contenedor
.env.example         Plantilla de configuración
```

## Notas de diseño

- **Destinatarios vs. segmentos facturados**: un mismo SMS puede facturarse
  como varios "segmentos" según su longitud. El panel usa siempre el número
  real de destinatarios para calcular efectividad, y reserva los segmentos
  facturados únicamente para referencia de costo.
- **Efectividad**: se calcula solo sobre destinatarios con resultado
  reportado (`destinatarios_evaluables`), para no diluir el promedio con
  campañas que aún no tienen resultados.
- **Impresión/PDF**: las tarjetas de color usan una paleta oscura alterna al
  imprimir (en vez de blanco sobre blanco), para que los gráficos se vean
  correctamente en el PDF exportado.
- **Entidades y dependencias**: nombres que corresponden a dependencias o
  áreas de una misma entidad (por ejemplo, las distintas formas de escribir
  la Secretaría General de la Alcaldía Mayor de Bogotá y sus dependencias
  como la Oficina Consejería Distrital de Paz, Víctimas y Reconciliación, o
  la Subsecretaría) se agrupan bajo un único nombre canónico, configurado en
  `etl._DEPENDENCIAS_SECRETARIA_GENERAL`.
- **Campañas con varias entidades**: cuando una celda de entidad menciona más
  de una (separadas por "/", p. ej. "Secretaría General /Prosperidad
  Social"), la campaña se cuenta en los rankings y filtros de ambas
  entidades por separado, sin perder la trazabilidad de que fue una
  solicitud conjunta.
