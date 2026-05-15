"""Validaciones del RMT: RUC, período, consistencia de totales."""
from __future__ import annotations

from dataclasses import dataclass


@dataclass
class ResultadoValidacion:
    ok: bool
    errores: list[str]


def validar_ruc(ruc: str) -> bool:
    """RUC Ecuador: 13 dígitos, los 3 últimos son '001' para persona natural / sociedad."""
    if not ruc or not ruc.isdigit() or len(ruc) != 13:
        return False
    return ruc[-3:] == "001"


def validar_periodo(periodo: str) -> bool:
    """Formato YYYY-MM."""
    if len(periodo) != 7 or periodo[4] != "-":
        return False
    y, m = periodo.split("-")
    return y.isdigit() and m.isdigit() and 1 <= int(m) <= 12
