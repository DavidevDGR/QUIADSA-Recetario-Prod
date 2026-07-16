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
# Servidor con instancia con nombre (SOLMICRO) y puerto NO estándar (1435):
# formato SERVER=IP\INSTANCIA,PUERTO
SOLMICRO_CONN_STRING = (
    "DRIVER={ODBC Driver 17 for SQL Server};"
    "SERVER=10.0.0.210\\SOLMICRO,1435;"
    "DATABASE=QUI_xQUIADSA;"
    "UID=lectura;"
    "PWD=lectura;"
)

# --- Consultas SQL a Solmicro (esquema real confirmado) ---
# La OF se busca por su código legible ("NOrden", el que se escanea por
# código de barras). Se filtran las secciones por IDTipoRuta (según lo
# confirmado), usando el campo "Texto" (ntext) como instrucción y
# "TiempoEjecUnit" directamente como minutos previstos de la sección.
SQL_ORDEN_FABRICACION = """
    SELECT
        NOrden,
        IDArticulo,
        IDTipoRuta,
        QFabricar,
        Lote
    FROM tbOrdenFabricacion
    WHERE NOrden = ?
"""

# El centro/máquina (IDCentro) sale de tbRuta, NO de tbOrdenFabricacion.
# Se toma el de la primera sección (Secuencia más baja) como centro de
# toda la OF, asumiendo que todas las operaciones de una misma ruta
# comparten centro. Si en la práctica varía entre secciones, avisa para
# ajustar el criterio.
SQL_SECCIONES_RUTA = """
    SELECT
        Secuencia,
        Texto,
        TiempoEjecUnit,
        IDCentro
    FROM tbRuta
    WHERE IDTipoRuta = ?
    ORDER BY Secuencia ASC
"""

# Identificador de esta máquina/puesto (si no se obtiene de la OF)
MAQUINA_ID_DEFECTO = "MAQUINA-01"

# --- Exportación a Excel / Nube ---
# Carpeta compartida en red donde se guarda el Excel de registro.
EXPORT_FOLDER = r"\\10.0.0.210\comun\FABRICA\AppFabricacion\exports"
EXCEL_FILENAME = "registro_fabricacion.xlsx"

# Usuario que realiza la exportación (aparece en el registro)
USUARIO_EXPORTACION = "fabricacion@quiadsa.com"

# --- Parámetros de negocio ---
# % de margen sobre el tiempo de ejecución de la sección antes de exigir motivo
MARGEN_TIEMPO_PORCENTAJE = 0.10  # 10%

# % de margen sobre la cantidad (kg) prevista en la OF, aplicado al peso neto
# final, antes de exigir motivo de incidencia
MARGEN_PESO_PORCENTAJE = 0.10  # 10%

# Comportamiento del botón Pausa/Reanudar:
# - False (por defecto, comportamiento actual): el contador de la sección
#   sigue corriendo en tiempo real aunque se pulse Pausa; la pausa solo
#   queda registrada (inicio/fin/minutos) para el Excel, pero no afecta al
#   cronómetro en pantalla ni al tiempo real de la sección.
# - True: al pulsar Pausa el contador se congela (p.ej. en 00:30); al
#   pulsar Reanudar continúa exactamente desde ese punto (00:31, 00:32...),
#   como si el tiempo pausado no hubiera existido. Esto también afecta al
#   tiempo real de la sección que se exporta al Excel y a la comprobación
#   del margen de tiempo (MARGEN_TIEMPO_PORCENTAJE).
PAUSA_DETIENE_CONTADOR = False

# Tamaño de fuente / interfaz táctil
FONT_FAMILY = "Segoe UI"
FONT_SIZE_NORMAL = 16
FONT_SIZE_TITLE = 24
FONT_SIZE_TIMER = 48
BTN_HEIGHT = 2
