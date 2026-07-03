# -*- coding: utf-8 -*-
"""
Capa de acceso a datos.

Contiene DOS partes claramente separadas:

1) SolmicroRepository
   -> Consultas de LECTURA al ERP Solmicro (Orden de Fabricación, Tipos de Ruta).
   -> AQUI está todo lo que debes adaptar a tu conexión real (pyodbc / cx_Oracle / etc).
   -> Por defecto incluye un modo "MOCK" (datos de prueba) para poder probar
      la aplicación sin tener la conexión real configurada.

2) LocalRepository
   -> Base de datos local (SQLite) donde se registra la ejecución de cada OF:
      operarios, secciones, incidencias, tiempos, etc.
   -> Esta información es la que luego se exporta a Excel.
"""

import sqlite3
import datetime
from dataclasses import dataclass, field
from typing import List, Optional

import config

# =========================================================================
# 1) SOLMICRO (ERP) - CONSULTAS DE LECTURA
# =========================================================================

# Cambia a False cuando tengas la conexión real a Solmicro configurada.
MOCK_MODE = True


@dataclass
class SeccionRuta:
    numero: int              # orden del paso (1..n)
    texto: str                # instrucción a mostrar
    tiempo_ejecucion_min: float  # tiempo previsto en minutos


@dataclass
class OrdenFabricacion:
    codigo_of: str
    codigo_tipo_ruta: str
    articulo: str
    cantidad_kg: float
    centro: str               # identificador de máquina
    secciones: List[SeccionRuta] = field(default_factory=list)


class SolmicroRepository:
    """Encapsula el acceso al ERP Solmicro. Adaptar aquí la conexión real."""

    def __init__(self):
        if not MOCK_MODE:
            # ---------------------------------------------------------------
            # EJEMPLO de conexión real con pyodbc (SQL Server). Descomentar
            # e instalar `pip install pyodbc` cuando se disponga de acceso.
            # ---------------------------------------------------------------
            # import pyodbc
            # self.conn = pyodbc.connect(config.SOLMICRO_CONN_STRING)
            raise NotImplementedError(
                "Configura la conexión real a Solmicro en db.py (SolmicroRepository.__init__)"
            )

    def obtener_orden_fabricacion(self, codigo_of: str) -> Optional[OrdenFabricacion]:
        """
        Dado el código de OF (leído por código de barras), devuelve:
        - código tipo de ruta
        - artículo (fórmula)
        - cantidad en kg a fabricar
        - centro (identificador de máquina)
        - lista de secciones (a través del tipo de ruta)
        """
        if MOCK_MODE:
            return self._mock_obtener_orden(codigo_of)

        # ---------------------------------------------------------------
        # EJEMPLO de consulta real. Ajustar nombres de tablas/columnas
        # a la instalación concreta de Solmicro.
        # ---------------------------------------------------------------
        # cursor = self.conn.cursor()
        # cursor.execute(
        #     """
        #     SELECT CodigoTipoRuta, Articulo, CantidadKg, Centro
        #     FROM OrdenesFabricacion
        #     WHERE CodigoOF = ?
        #     """,
        #     codigo_of,
        # )
        # row = cursor.fetchone()
        # if not row:
        #     return None
        # of_data = OrdenFabricacion(
        #     codigo_of=codigo_of,
        #     codigo_tipo_ruta=row.CodigoTipoRuta,
        #     articulo=row.Articulo,
        #     cantidad_kg=row.CantidadKg,
        #     centro=row.Centro,
        # )
        # of_data.secciones = self._obtener_secciones(of_data.codigo_tipo_ruta)
        # return of_data
        raise NotImplementedError

    def _obtener_secciones(self, codigo_tipo_ruta: str) -> List[SeccionRuta]:
        if MOCK_MODE:
            return self._mock_secciones(codigo_tipo_ruta)

        # ---------------------------------------------------------------
        # cursor = self.conn.cursor()
        # cursor.execute(
        #     """
        #     SELECT NumeroSeccion, Texto, TiempoEjecucionMin
        #     FROM TiposRuta
        #     WHERE CodigoTipoRuta = ?
        #     ORDER BY NumeroSeccion ASC
        #     """,
        #     codigo_tipo_ruta,
        # )
        # return [
        #     SeccionRuta(numero=r.NumeroSeccion, texto=r.Texto,
        #                  tiempo_ejecucion_min=r.TiempoEjecucionMin)
        #     for r in cursor.fetchall()
        # ]
        raise NotImplementedError

    # ------------------------------------------------------------------
    # DATOS DE PRUEBA (MOCK) - permiten ejecutar la app sin Solmicro real
    # ------------------------------------------------------------------
    def _mock_obtener_orden(self, codigo_of: str) -> Optional[OrdenFabricacion]:
        # Cualquier código de OF que empiece por "OF" devuelve datos de ejemplo.
        if not codigo_of:
            return None
        of_data = OrdenFabricacion(
            codigo_of=codigo_of,
            codigo_tipo_ruta="RUTA-STD-01",
            articulo="FORMULA-DEMO-A",
            cantidad_kg=500.0,
            centro=config.MAQUINA_ID_DEFECTO,
        )
        of_data.secciones = self._mock_secciones(of_data.codigo_tipo_ruta)
        return of_data

    def _mock_secciones(self, codigo_tipo_ruta: str) -> List[SeccionRuta]:
        return [
            SeccionRuta(1, "Cargar agua y activar agitación a velocidad baja.", 1),
            SeccionRuta(2, "Añadir materia prima A poco a poco.", 1),
            SeccionRuta(3, "Subir a velocidad media y mantener homogeneización.", 1),
            SeccionRuta(4, "Añadir materia prima B y controlar temperatura.", 1),
            SeccionRuta(5, "Enfriar y comprobar viscosidad final.", 1),
        ]


# =========================================================================
# 2) BASE DE DATOS LOCAL - REGISTRO DE EJECUCIÓN
# =========================================================================

class LocalRepository:
    """Gestiona el histórico local de ejecuciones de OF (SQLite)."""

    def __init__(self, db_path: str = None):
        self.db_path = db_path or config.LOCAL_DB_PATH
        self._crear_esquema()

    def _conn(self):
        conn = sqlite3.connect(self.db_path)
        conn.execute("PRAGMA foreign_keys = ON")
        return conn

    def _crear_esquema(self):
        with self._conn() as c:
            c.execute("""
                CREATE TABLE IF NOT EXISTS ejecucion_of (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    codigo_of TEXT NOT NULL,
                    maquina_id TEXT NOT NULL,
                    operarios TEXT NOT NULL,        -- lista separada por comas
                    cuba_id TEXT NOT NULL,
                    peso_tara REAL NOT NULL,
                    fecha_inicio TEXT NOT NULL,
                    fecha_fin TEXT,
                    estado TEXT NOT NULL,           -- EN_CURSO / FINALIZADA / CANCELADA
                    peso_llena REAL,
                    tiempo_total_min REAL
                )
            """)
            c.execute("""
                CREATE TABLE IF NOT EXISTS ejecucion_seccion (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    ejecucion_of_id INTEGER NOT NULL,
                    numero_seccion INTEGER NOT NULL,
                    texto TEXT NOT NULL,
                    tiempo_previsto_min REAL NOT NULL,
                    fecha_inicio TEXT NOT NULL,
                    fecha_fin TEXT,
                    tiempo_real_min REAL,
                    incidencia_motivo TEXT,
                    FOREIGN KEY (ejecucion_of_id) REFERENCES ejecucion_of(id)
                )
            """)

    # ---------------- Ejecución de OF ----------------

    def crear_ejecucion(self, codigo_of, maquina_id, operarios: List[str],
                         cuba_id, peso_tara) -> int:
        with self._conn() as c:
            cur = c.execute("""
                INSERT INTO ejecucion_of
                    (codigo_of, maquina_id, operarios, cuba_id, peso_tara,
                     fecha_inicio, estado)
                VALUES (?, ?, ?, ?, ?, ?, 'EN_CURSO')
            """, (
                codigo_of, maquina_id, ",".join(operarios), cuba_id, peso_tara,
                datetime.datetime.now().isoformat(timespec="seconds"),
            ))
            return cur.lastrowid

    def finalizar_ejecucion(self, ejecucion_id: int, peso_llena: float):
        with self._conn() as c:
            fecha_inicio = c.execute(
                "SELECT fecha_inicio FROM ejecucion_of WHERE id=?", (ejecucion_id,)
            ).fetchone()[0]
            fecha_fin = datetime.datetime.now()
            inicio = datetime.datetime.fromisoformat(fecha_inicio)
            tiempo_total = (fecha_fin - inicio).total_seconds() / 60.0
            c.execute("""
                UPDATE ejecucion_of
                SET estado='FINALIZADA', fecha_fin=?, peso_llena=?, tiempo_total_min=?
                WHERE id=?
            """, (fecha_fin.isoformat(timespec="seconds"), peso_llena,
                  round(tiempo_total, 2), ejecucion_id))

    def cancelar_ejecucion(self, ejecucion_id: int):
        with self._conn() as c:
            fecha_inicio = c.execute(
                "SELECT fecha_inicio FROM ejecucion_of WHERE id=?", (ejecucion_id,)
            ).fetchone()[0]
            fecha_fin = datetime.datetime.now()
            inicio = datetime.datetime.fromisoformat(fecha_inicio)
            tiempo_total = (fecha_fin - inicio).total_seconds() / 60.0
            c.execute("""
                UPDATE ejecucion_of
                SET estado='CANCELADA', fecha_fin=?, tiempo_total_min=?
                WHERE id=?
            """, (fecha_fin.isoformat(timespec="seconds"),
                  round(tiempo_total, 2), ejecucion_id))

    # ---------------- Secciones ----------------

    def iniciar_seccion(self, ejecucion_of_id, numero_seccion, texto,
                         tiempo_previsto_min) -> int:
        with self._conn() as c:
            cur = c.execute("""
                INSERT INTO ejecucion_seccion
                    (ejecucion_of_id, numero_seccion, texto, tiempo_previsto_min,
                     fecha_inicio)
                VALUES (?, ?, ?, ?, ?)
            """, (ejecucion_of_id, numero_seccion, texto, tiempo_previsto_min,
                  datetime.datetime.now().isoformat(timespec="seconds")))
            return cur.lastrowid

    def cerrar_seccion(self, seccion_id: int, incidencia_motivo: Optional[str] = None):
        with self._conn() as c:
            fecha_inicio = c.execute(
                "SELECT fecha_inicio FROM ejecucion_seccion WHERE id=?", (seccion_id,)
            ).fetchone()[0]
            fecha_fin = datetime.datetime.now()
            inicio = datetime.datetime.fromisoformat(fecha_inicio)
            tiempo_real = (fecha_fin - inicio).total_seconds() / 60.0
            c.execute("""
                UPDATE ejecucion_seccion
                SET fecha_fin=?, tiempo_real_min=?, incidencia_motivo=?
                WHERE id=?
            """, (fecha_fin.isoformat(timespec="seconds"), round(tiempo_real, 2),
                  incidencia_motivo, seccion_id))

    # ---------------- Consultas para exportación ----------------

    def obtener_ejecucion(self, ejecucion_id: int) -> dict:
        with self._conn() as c:
            c.row_factory = sqlite3.Row
            row = c.execute(
                "SELECT * FROM ejecucion_of WHERE id=?", (ejecucion_id,)
            ).fetchone()
            return dict(row) if row else None

    def obtener_secciones(self, ejecucion_id: int) -> List[dict]:
        with self._conn() as c:
            c.row_factory = sqlite3.Row
            rows = c.execute("""
                SELECT * FROM ejecucion_seccion
                WHERE ejecucion_of_id=?
                ORDER BY numero_seccion ASC
            """, (ejecucion_id,)).fetchall()
            return [dict(r) for r in rows]
