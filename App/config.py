# -*- coding: utf-8 -*-
"""
Configuración general de la aplicación.
Ajusta estos valores según tu instalación.
"""

import os
import sys

# ---------------------------------------------------------------------
# Carpeta base de la aplicación.
# - En desarrollo (python main.py): la carpeta donde está este fichero.
# - Empaquetada con PyInstaller (.exe): la carpeta donde está el .exe,
#   NO la carpeta temporal (_MEIxxxxx) donde PyInstaller descomprime los
#   recursos en modo "un solo fichero". Si no se hiciera así, la base de
#   datos y el Excel se borrarían cada vez que se cierra el programa.
# ---------------------------------------------------------------------
if getattr(sys, "frozen", False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# --- Base de datos local (registro de ejecución de OFs) ---
LOCAL_DB_PATH = os.path.join(BASE_DIR, "cuba_app.db")

# --- Conexión a Solmicro (ERP) ---
# TODO: sustituir por los datos reales de conexión (SQL Server).
# Pídele a IT/soporte de Solmicro estos datos (ver checklist en README.md):
#   - Servidor (IP o nombre) y puerto (si no es el 1433 por defecto)
#   - Nombre de la base de datos
#   - Usuario y contraseña (a ser posible, un usuario de solo lectura)
#   - Driver ODBC instalado en el PC (ej. "ODBC Driver 17 for SQL Server")
SOLMICRO_CONN_STRING = (
    "DRIVER={ODBC Driver 17 for SQL Server};"
    "SERVER=SERVIDOR_SOLMICRO;"
    "DATABASE=Solmicro;"
    "UID=usuario;"
    "PWD=clave;"
)

# --- Consultas SQL a Solmicro ---
# TODO: sustituir estos nombres de tabla/columna por los reales de tu
# instalación. Usa `explorar_solmicro.py` (incluido en este proyecto) para
# localizarlos sin tener que adivinarlos: busca tablas relacionadas con
# "orden", "fabrica", "ruta", "seccion", etc. y sus columnas.
#
# La consulta de la OF debe devolver, para el código de OF dado (parámetro
# "?"), estas 4 columnas con esos alias exactos (o renombra con AS):
#   CodigoTipoRuta, Articulo, CantidadKg, Centro
SQL_ORDEN_FABRICACION = """
    SELECT
        CodigoTipoRuta,
        Articulo,
        CantidadKg,
        Centro
    FROM OrdenesFabricacion   -- <-- AJUSTAR nombre real de la tabla
    WHERE CodigoOF = ?        -- <-- AJUSTAR nombre real de la columna
"""

# La consulta de secciones debe devolver, para el código de tipo de ruta
# dado (parámetro "?"), estas 3 columnas con esos alias exactos, ordenadas
# por número de sección ascendente:
#   NumeroSeccion, Texto, TiempoEjecucionMin
SQL_SECCIONES_RUTA = """
    SELECT
        NumeroSeccion,
        Texto,
        TiempoEjecucionMin
    FROM TiposRuta                    -- <-- AJUSTAR nombre real de la tabla
    WHERE CodigoTipoRuta = ?          -- <-- AJUSTAR nombre real de la columna
    ORDER BY NumeroSeccion ASC
"""

# Identificador de esta máquina/puesto (si no se obtiene de la OF)
MAQUINA_ID_DEFECTO = "CUBA-01"

# --- Exportación a Excel / Nube ---
# Carpeta local de trabajo antes de subir a la nube (OneDrive/Google Drive
# sincronizado como carpeta local es la forma más simple de "guardar en la nube").
EXPORT_FOLDER = os.path.join(BASE_DIR, "exports")
EXCEL_FILENAME = "registro_fabricacion.xlsx"

# Usuario que realiza la exportación (aparece en el registro)
USUARIO_EXPORTACION = "fabricacion@quiadsa.com"

# --- Parámetros de negocio ---
# % de margen sobre el tiempo de ejecución de la sección antes de exigir motivo
MARGEN_TIEMPO_PORCENTAJE = 0.10  # 10%

# % de margen sobre la cantidad (kg) prevista en la OF, aplicado al peso neto
# final, antes de exigir motivo de incidencia
MARGEN_PESO_PORCENTAJE = 0.10  # 10%

# Tamaño de fuente / interfaz táctil
FONT_FAMILY = "Segoe UI"
FONT_SIZE_NORMAL = 16
FONT_SIZE_TITLE = 24
FONT_SIZE_TIMER = 48
BTN_HEIGHT = 2
