"""Tests para modelos de datos."""
from pathlib import Path

from src.models import DocumentoRMT, RMTProcesado


def _rmt(inicio: str, fin: str) -> RMTProcesado:
    return RMTProcesado(
        archivo=Path("dummy.xlsx"),
        nombre_archivo="dummy.xlsx",
        hash_archivo="x",
        ruc="1793205060001",
        cliente_nombre="DEMO S.A.",
        periodo_inicio=inicio,
        periodo_fin=fin,
    )


def test_periodo_label_mes_unico():
    assert _rmt("2026-04", "2026-04").periodo_label == "2026-04"


def test_periodo_label_rango():
    assert _rmt("2026-04", "2026-05").periodo_label == "2026-04 a 2026-05"


def test_documento_propiedades_netas():
    doc = DocumentoRMT(
        bloque="FACTURAS EMITIDAS",
        grupo="venta",
        numero="001-001-001",
        fecha_emision=None,
        fecha_carga=None,
        contraparte_id="0912345678001",
        contraparte_nombre="CLIENTE",
        base_0=100.0,
        descuento_0=10.0,
        base_iva=200.0,
        descuento_iva=20.0,
        no_objeto_iva=50.0,
        descuento_no_objeto=5.0,
        servicio=10.0,
    )
    assert doc.base_0_neta == 90.0
    assert doc.base_iva_neta == 180.0
    assert doc.no_objeto_neto == 55.0  # 50 - 5 + 10
