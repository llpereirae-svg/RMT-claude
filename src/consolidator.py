"""Consultas consolidadas sobre las bases SQLite locales."""
from __future__ import annotations

from pathlib import Path

import pandas as pd

from . import db


def consolidar(data_root: Path, ruc: str | None = None, periodo: str | None = None) -> pd.DataFrame:
    frames: list[pd.DataFrame] = []
    for db_file in data_root.glob("*/*.db"):
        query = """
            SELECT a.ruc, a.cliente_nombre, a.periodo_inicio, a.periodo_fin, a.nombre AS archivo,
                   d.grupo, d.bloque, d.numero, d.fecha_emision, d.contraparte_id,
                   d.contraparte_nombre,
                   ROUND(d.base_0 - d.descuento_0, 2) AS base_0,
                   ROUND(d.base_iva - d.descuento_iva, 2) AS base_iva,
                   ROUND(d.no_objeto_iva - d.descuento_no_objeto + d.servicio, 2) AS no_objeto,
                   d.iva, d.total
            FROM documentos d
            JOIN archivos a ON a.id = d.archivo_id
            WHERE a.estado = 'ok'
        """
        params: list[str] = []
        if ruc:
            query += " AND a.ruc = ?"
            params.append(ruc)
        if periodo:
            query += " AND (a.periodo_inicio = ? OR a.periodo_fin = ?)"
            params.extend([periodo, periodo])
        with db.get_conn(db_file) as conn:
            frames.append(pd.read_sql_query(query, conn, params=params))
    if not frames:
        return pd.DataFrame()
    return pd.concat(frames, ignore_index=True)


def exportar_excel(df: pd.DataFrame, destino: Path) -> Path:
    destino.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(destino, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Consolidado")
    return destino
