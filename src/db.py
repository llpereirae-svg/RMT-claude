"""SQLite local por cliente/RUC."""
from __future__ import annotations

import json
import sqlite3
from contextlib import contextmanager
from dataclasses import asdict
from datetime import date, datetime
from pathlib import Path
from typing import Iterable

from .models import FormularioCalculado, RMTProcesado

ROOT = Path(__file__).resolve().parent.parent
DEFAULT_DATA_ROOT = ROOT / "data" / "clientes"

SCHEMA = """
PRAGMA foreign_keys = ON;

CREATE TABLE IF NOT EXISTS archivos (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    nombre TEXT NOT NULL,
    ruc TEXT NOT NULL,
    cliente_nombre TEXT NOT NULL,
    periodo_inicio TEXT NOT NULL,
    periodo_fin TEXT NOT NULL,
    hash_archivo TEXT NOT NULL UNIQUE,
    estado TEXT NOT NULL,
    mensaje_error TEXT,
    creado_en TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS documentos (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    archivo_id INTEGER NOT NULL REFERENCES archivos(id) ON DELETE CASCADE,
    bloque TEXT NOT NULL,
    grupo TEXT NOT NULL,
    numero TEXT,
    fecha_emision TEXT,
    fecha_carga TEXT,
    contraparte_id TEXT,
    contraparte_nombre TEXT,
    base_0 REAL DEFAULT 0,
    descuento_0 REAL DEFAULT 0,
    base_iva REAL DEFAULT 0,
    descuento_iva REAL DEFAULT 0,
    no_objeto_iva REAL DEFAULT 0,
    descuento_no_objeto REAL DEFAULT 0,
    total_sin_impuesto REAL DEFAULT 0,
    iva REAL DEFAULT 0,
    servicio REAL DEFAULT 0,
    propina REAL DEFAULT 0,
    total REAL DEFAULT 0,
    sustento TEXT,
    autorizacion TEXT,
    doc_modificado TEXT,
    estado TEXT,
    cuadre REAL DEFAULT 0,
    extra_json TEXT
);

CREATE TABLE IF NOT EXISTS retenciones (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    archivo_id INTEGER NOT NULL REFERENCES archivos(id) ON DELETE CASCADE,
    bloque TEXT NOT NULL,
    numero TEXT,
    fecha_emision TEXT,
    contraparte_id TEXT,
    contraparte_nombre TEXT,
    factura TEXT,
    iva_10 REAL DEFAULT 0,
    iva_20 REAL DEFAULT 0,
    iva_30 REAL DEFAULT 0,
    iva_50 REAL DEFAULT 0,
    iva_70 REAL DEFAULT 0,
    iva_100 REAL DEFAULT 0,
    codigo_ir TEXT,
    base_ir REAL DEFAULT 0,
    valor_ir REAL DEFAULT 0,
    extra_json TEXT
);

CREATE TABLE IF NOT EXISTS anulados (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    archivo_id INTEGER NOT NULL REFERENCES archivos(id) ON DELETE CASCADE,
    tipo TEXT,
    fecha TEXT,
    numero TEXT,
    autorizacion TEXT,
    contraparte_id TEXT,
    contraparte_nombre TEXT,
    total REAL DEFAULT 0
);

CREATE TABLE IF NOT EXISTS casilleros (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    archivo_id INTEGER NOT NULL REFERENCES archivos(id) ON DELETE CASCADE,
    formulario TEXT NOT NULL,
    casillero TEXT NOT NULL,
    valor REAL NOT NULL
);

CREATE TABLE IF NOT EXISTS advertencias (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    archivo_id INTEGER NOT NULL REFERENCES archivos(id) ON DELETE CASCADE,
    mensaje TEXT NOT NULL
);

CREATE INDEX IF NOT EXISTS idx_archivos_ruc_periodo ON archivos(ruc, periodo_inicio, periodo_fin);
CREATE INDEX IF NOT EXISTS idx_documentos_archivo_grupo ON documentos(archivo_id, grupo);
CREATE INDEX IF NOT EXISTS idx_documentos_contraparte ON documentos(contraparte_id, contraparte_nombre);
CREATE INDEX IF NOT EXISTS idx_retenciones_archivo ON retenciones(archivo_id);
CREATE INDEX IF NOT EXISTS idx_casilleros_archivo_form ON casilleros(archivo_id, formulario);
"""


def safe_name(value: str) -> str:
    cleaned = "".join(ch if ch.isalnum() or ch in ("-", "_") else "_" for ch in value.strip())
    return cleaned.strip("_") or "cliente"


def client_db_path(ruc: str, cliente_nombre: str = "", data_root: Path = DEFAULT_DATA_ROOT) -> Path:
    folder = data_root / f"{safe_name(ruc)}_{safe_name(cliente_nombre)[:48]}"
    return folder / "rmt.db"


def init_db(db_path: Path) -> None:
    db_path.parent.mkdir(parents=True, exist_ok=True)
    with sqlite3.connect(db_path) as conn:
        conn.executescript(SCHEMA)


@contextmanager
def get_conn(db_path: Path):
    init_db(db_path)
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    # PRAGMA foreign_keys es por-conexion en SQLite. Sin esto, ON DELETE CASCADE
    # no se aplica y reprocesar un RMT deja documentos/casilleros huerfanos.
    conn.execute("PRAGMA foreign_keys = ON")
    try:
        yield conn
        conn.commit()
    finally:
        conn.close()


def save_result(db_path: Path, rmt: RMTProcesado, formularios: FormularioCalculado) -> int:
    with get_conn(db_path) as conn:
        conn.execute("DELETE FROM archivos WHERE hash_archivo = ?", (rmt.hash_archivo,))
        cur = conn.execute(
            """
            INSERT INTO archivos
                (nombre, ruc, cliente_nombre, periodo_inicio, periodo_fin, hash_archivo, estado, mensaje_error)
            VALUES (?, ?, ?, ?, ?, ?, 'ok', NULL)
            """,
            (
                rmt.nombre_archivo,
                rmt.ruc,
                rmt.cliente_nombre,
                rmt.periodo_inicio,
                rmt.periodo_fin,
                rmt.hash_archivo,
            ),
        )
        archivo_id = int(cur.lastrowid)
        _insert_many(conn, "documentos", archivo_id, rmt.documentos)
        _insert_many(conn, "retenciones", archivo_id, rmt.retenciones)
        _insert_many(conn, "anulados", archivo_id, rmt.anulados)
        for form, casilleros in (("104", formularios.casilleros_104), ("103", formularios.casilleros_103)):
            conn.executemany(
                "INSERT INTO casilleros (archivo_id, formulario, casillero, valor) VALUES (?, ?, ?, ?)",
                [(archivo_id, form, key, value) for key, value in sorted(casilleros.items()) if value],
            )
        conn.executemany(
            "INSERT INTO advertencias (archivo_id, mensaje) VALUES (?, ?)",
            [(archivo_id, msg) for msg in formularios.advertencias],
        )
        return archivo_id


def latest_files(data_root: Path = DEFAULT_DATA_ROOT, limit: int = 200) -> list[dict]:
    rows: list[dict] = []
    for db_file in data_root.glob("*/*.db"):
        with get_conn(db_file) as conn:
            for row in conn.execute(
                """
                SELECT id, nombre, ruc, cliente_nombre, periodo_inicio, periodo_fin, estado, creado_en
                FROM archivos
                ORDER BY creado_en DESC
                LIMIT ?
                """,
                (limit,),
            ):
                item = dict(row)
                item["db_path"] = str(db_file)
                rows.append(item)
    return sorted(rows, key=lambda x: x["creado_en"], reverse=True)[:limit]


def casilleros(db_path: Path, archivo_id: int, formulario: str) -> list[dict]:
    with get_conn(db_path) as conn:
        return [
            dict(row)
            for row in conn.execute(
                """
                SELECT casillero, valor
                FROM casilleros
                WHERE archivo_id = ? AND formulario = ?
                ORDER BY CAST(casillero AS INTEGER), casillero
                """,
                (archivo_id, formulario),
            )
        ]


def documentos_resumen(db_path: Path, archivo_id: int) -> list[dict]:
    with get_conn(db_path) as conn:
        return [
            dict(row)
            for row in conn.execute(
                """
                SELECT grupo, bloque, COUNT(*) documentos,
                       ROUND(SUM(base_0 - descuento_0), 2) base_0,
                       ROUND(SUM(base_iva - descuento_iva), 2) base_iva,
                       ROUND(SUM(no_objeto_iva), 2) no_objeto,
                       ROUND(SUM(iva), 2) iva,
                       ROUND(SUM(total), 2) total
                FROM documentos
                WHERE archivo_id = ?
                GROUP BY grupo, bloque
                ORDER BY grupo, bloque
                """,
                (archivo_id,),
            )
        ]


def _insert_many(conn: sqlite3.Connection, table: str, archivo_id: int, items: Iterable[object]) -> None:
    for item in items:
        data = asdict(item)
        data["archivo_id"] = archivo_id
        for key, value in list(data.items()):
            if isinstance(value, (date, datetime)):
                data[key] = value.isoformat()
            elif isinstance(value, dict):
                data[f"{key}_json"] = json.dumps(value, ensure_ascii=False)
                del data[key]
        columns = list(data)
        placeholders = ", ".join("?" for _ in columns)
        conn.execute(
            f"INSERT INTO {table} ({', '.join(columns)}) VALUES ({placeholders})",
            [data[col] for col in columns],
        )
