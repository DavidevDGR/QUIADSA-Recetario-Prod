# -*- coding: utf-8 -*-
"""
Configuración general de la aplicación.
Ajusta estos valores según tu instalación.
"""

# --- Base de datos local (registro de ejecución de OFs) ---
LOCAL_DB_PATH = "cuba_app.db"

# --- Conexión a Solmicro (ERP) ---
# TODO: sustituir por los datos reales de conexión (SQL Server normalmente).
# Ejemplo típico con pyodbc:
#   DRIVER={ODBC Driver 17 for SQL Server};SERVER=IP_SERVIDOR;DATABASE=Solmicro;UID=usuario;PWD=clave
SOLMICRO_CONN_STRING = (
    "DRIVER={ODBC Driver 17 for SQL Server};"
    "SERVER=SERVIDOR_SOLMICRO;"
    "DATABASE=Solmicro;"
    "UID=usuario;"
    "PWD=clave;"
)

# Identificador de esta máquina/puesto (si no se obtiene de la OF)
MAQUINA_ID_DEFECTO = "CUBA-01"

# --- Exportación a Excel / Nube ---
# Carpeta local de trabajo antes de subir a la nube (OneDrive/Google Drive
# sincronizado como carpeta local es la forma más simple de "guardar en la nube").
EXPORT_FOLDER = "exports"
EXCEL_FILENAME = "registro_fabricacion.xlsx"

# Usuario que realiza la exportación (aparece en el registro)
USUARIO_EXPORTACION = "fabricacion@quiadsa.com"

# --- Parámetros de negocio ---
# % de margen sobre el tiempo de ejecución de la sección antes de exigir motivo
MARGEN_TIEMPO_PORCENTAJE = 0.25  # 25%

# Tamaño de fuente / interfaz táctil
FONT_FAMILY = "Segoe UI"
FONT_SIZE_NORMAL = 16
FONT_SIZE_TITLE = 24
FONT_SIZE_TIMER = 48
BTN_HEIGHT = 2
