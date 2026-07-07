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
      operarios (iniciales y por sección), secciones, incidencias (tiempo y
      peso), pausas registradas, tiempos, etc.
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
MOCK_MODE = False


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
    """Encapsula el acceso al ERP Solmicro. En modo real (MOCK_MODE=False)
    abre una conexión de solo lectura por pyodbc usando la cadena de
    conexión y las consultas SQL definidas en config.py."""

    def __init__(self):
        self.conn = None
        if not MOCK_MODE:
            import pyodbc  # requiere: pip install pyodbc
            self.conn = pyodbc.connect(config.SOLMICRO_CONN_STRING, timeout=5)

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

        cursor = self.conn.cursor()
        cursor.execute(config.SQL_ORDEN_FABRICACION, codigo_of)
        row = cursor.fetchone()
        if not row:
            return None

        of_data = OrdenFabricacion(
            codigo_of=str(row.NOrden),
            codigo_tipo_ruta=str(row.IDTipoRuta) if row.IDTipoRuta else "",
            articulo=str(row.IDArticulo),
            cantidad_kg=float(row.QFabricar),
            centro=str(row.IDCentroGestion),
        )
        of_data.secciones = self._obtener_secciones(of_data.codigo_tipo_ruta)
        return of_data

    def _obtener_secciones(self, codigo_tipo_ruta: str) -> List[SeccionRuta]:
        if MOCK_MODE:
            return self._mock_secciones(codigo_tipo_ruta)

        cursor = self.conn.cursor()
        cursor.execute(config.SQL_SECCIONES_RUTA, codigo_tipo_ruta)
        return [
            SeccionRuta(
                numero=int(r.Secuencia),
                texto=str(r.Texto) if r.Texto else "",
                tiempo_ejecucion_min=float(r.TiempoEjecUnit),
            )
            for r in cursor.fetchall()
        ]

    # ------------------------------------------------------------------
    # DATOS DE PRUEBA (MOCK) - permiten ejecutar la app sin Solmicro real
    # ------------------------------------------------------------------
    def _mock_obtener_orden(self, codigo_of: str) -> Optional[OrdenFabricacion]:
        # Cualquier código de OF no vacío devuelve datos de ejemplo.
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
            SeccionRuta(1, "Cargar agua y activar agitación a velocidad baja.", 5),
            SeccionRuta(2, "Añadir materia prima A poco a poco.", 8),
            SeccionRuta(3, "Subir a velocidad media y mantener homogeneización.", 10),
            SeccionRuta(4, "Añadir materia prima B y controlar temperatura.", 6),
            SeccionRuta(5, "Enfriar y comprobar viscosidad final.", 12),
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
                    operarios TEXT NOT NULL,        -- operarios iniciales, separados por comas
                    cuba_id TEXT NOT NULL,
                    peso_tara REAL NOT NULL,
                    cantidad_of_kg REAL,             -- cantidad prevista en la OF (kg)
                    fecha_inicio TEXT NOT NULL,
                    fecha_fin TEXT,
                    estado TEXT NOT NULL,           -- EN_CURSO / FINALIZADA / CANCELADA
                    peso_llena REAL,
                    tiempo_total_min REAL,
                    peso_incidencia_motivo TEXT
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
            c.execute("""
                CREATE TABLE IF NOT EXISTS ejecucion_seccion_operario (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    ejecucion_seccion_id INTEGER NOT NULL,
                    operario TEXT NOT NULL,
                    FOREIGN KEY (ejecucion_seccion_id) REFERENCES ejecucion_seccion(id)
                )
            """)
            c.execute("""
                CREATE TABLE IF NOT EXISTS ejecucion_pausa (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    ejecucion_of_id INTEGER NOT NULL,
                    ejecucion_seccion_id INTEGER,
                    fecha_inicio TEXT NOT NULL,
                    fecha_fin TEXT,
                    tiempo_total_min REAL,
                    FOREIGN KEY (ejecucion_of_id) REFERENCES ejecucion_of(id),
                    FOREIGN KEY (ejecucion_seccion_id) REFERENCES ejecucion_seccion(id)
                )
            """)
            # Migración suave por si la BD ya existía sin las columnas nuevas
            cols_of = {row[1] for row in c.execute("PRAGMA table_info(ejecucion_of)")}
            if "cantidad_of_kg" not in cols_of:
                c.execute("ALTER TABLE ejecucion_of ADD COLUMN cantidad_of_kg REAL")
            if "peso_incidencia_motivo" not in cols_of:
                c.execute("ALTER TABLE ejecucion_of ADD COLUMN peso_incidencia_motivo TEXT")
            cols_pausa = {row[1] for row in c.execute("PRAGMA table_info(ejecucion_pausa)")}
            if "ejecucion_seccion_id" not in cols_pausa:
                c.execute("ALTER TABLE ejecucion_pausa ADD COLUMN ejecucion_seccion_id INTEGER")

    # ---------------- Ejecución de OF ----------------

    def crear_ejecucion(self, codigo_of, maquina_id, operarios: List[str],
                         cuba_id, peso_tara, cantidad_of_kg: float = None) -> int:
        with self._conn() as c:
            cur = c.execute("""
                INSERT INTO ejecucion_of
                    (codigo_of, maquina_id, operarios, cuba_id, peso_tara,
                     cantidad_of_kg, fecha_inicio, estado)
                VALUES (?, ?, ?, ?, ?, ?, ?, 'EN_CURSO')
            """, (
                codigo_of, maquina_id, ",".join(operarios), cuba_id, peso_tara,
                cantidad_of_kg,
                datetime.datetime.now().isoformat(timespec="seconds"),
            ))
            return cur.lastrowid

    def finalizar_ejecucion(self, ejecucion_id: int, peso_llena: float,
                             peso_incidencia_motivo: Optional[str] = None):
        with self._conn() as c:
            fecha_inicio = c.execute(
                "SELECT fecha_inicio FROM ejecucion_of WHERE id=?", (ejecucion_id,)
            ).fetchone()[0]
            fecha_fin = datetime.datetime.now()
            inicio = datetime.datetime.fromisoformat(fecha_inicio)
            tiempo_total = (fecha_fin - inicio).total_seconds() / 60.0
            c.execute("""
                UPDATE ejecucion_of
                SET estado='FINALIZADA', fecha_fin=?, peso_llena=?, tiempo_total_min=?,
                    peso_incidencia_motivo=?
                WHERE id=?
            """, (fecha_fin.isoformat(timespec="seconds"), peso_llena,
                  round(tiempo_total, 2), peso_incidencia_motivo, ejecucion_id))

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

    # ---------------- Operarios por sección ----------------

    def registrar_operarios_seccion(self, seccion_id: int, operarios: List[str]):
        """Guarda una 'foto' de qué operarios estaban registrados en el
        momento de cerrar esta sección (para trazabilidad y exportación)."""
        with self._conn() as c:
            for op in operarios:
                c.execute("""
                    INSERT INTO ejecucion_seccion_operario (ejecucion_seccion_id, operario)
                    VALUES (?, ?)
                """, (seccion_id, op))

    def obtener_operarios_seccion(self, seccion_id: int) -> List[str]:
        with self._conn() as c:
            rows = c.execute("""
                SELECT operario FROM ejecucion_seccion_operario
                WHERE ejecucion_seccion_id=?
                ORDER BY id ASC
            """, (seccion_id,)).fetchall()
            return [r[0] for r in rows]

    # ---------------- Pausas ----------------

    def iniciar_pausa(self, ejecucion_of_id: int,
                       ejecucion_seccion_id: Optional[int] = None) -> int:
        with self._conn() as c:
            cur = c.execute("""
                INSERT INTO ejecucion_pausa (ejecucion_of_id, ejecucion_seccion_id, fecha_inicio)
                VALUES (?, ?, ?)
            """, (ejecucion_of_id, ejecucion_seccion_id,
                  datetime.datetime.now().isoformat(timespec="seconds")))
            return cur.lastrowid

    def finalizar_pausa(self, pausa_id: int):
        with self._conn() as c:
            fecha_inicio = c.execute(
                "SELECT fecha_inicio FROM ejecucion_pausa WHERE id=?", (pausa_id,)
            ).fetchone()[0]
            fecha_fin = datetime.datetime.now()
            inicio = datetime.datetime.fromisoformat(fecha_inicio)
            tiempo_total = (fecha_fin - inicio).total_seconds() / 60.0
            c.execute("""
                UPDATE ejecucion_pausa
                SET fecha_fin=?, tiempo_total_min=?
                WHERE id=?
            """, (fecha_fin.isoformat(timespec="seconds"), round(tiempo_total, 2), pausa_id))

    def obtener_pausas(self, ejecucion_of_id: int) -> List[dict]:
        with self._conn() as c:
            c.row_factory = sqlite3.Row
            rows = c.execute("""
                SELECT * FROM ejecucion_pausa
                WHERE ejecucion_of_id=?
                ORDER BY id ASC
            """, (ejecucion_of_id,)).fetchall()
            return [dict(r) for r in rows]

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
            secciones = [dict(r) for r in rows]
        # Añadir la lista de operarios registrados en cada sección
        for sec in secciones:
            sec["operarios"] = self.obtener_operarios_seccion(sec["id"])
        return secciones
