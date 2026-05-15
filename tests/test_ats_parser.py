"""Tests del parser ATS XML."""
from pathlib import Path

import pytest

from src.ats_parser import parse_ats


def _write_xml(tmp_path: Path, content: str) -> Path:
    f = tmp_path / "ats.xml"
    f.write_text(content, encoding="utf-8")
    return f


def test_parse_cabecera_y_estructura_basica(tmp_path: Path):
    xml = """<?xml version="1.0" encoding="UTF-8"?>
<iva>
    <TipoIDInformante>04</TipoIDInformante>
    <IdInformante>1793205060001</IdInformante>
    <razonSocial>DEMO S.A.</razonSocial>
    <Anio>2026</Anio>
    <Mes>04</Mes>
    <totalVentas>1000.00</totalVentas>
    <compras/>
    <ventas/>
    <exportaciones/>
    <anulados/>
</iva>"""
    path = _write_xml(tmp_path, xml)
    ats = parse_ats(path)
    assert ats.ruc_informante == "1793205060001"
    assert ats.razon_social == "DEMO S.A."
    assert ats.anio == "2026"
    assert ats.mes == "04"
    assert ats.total_ventas == 1000.00
    assert ats.compras == []
    assert ats.exportaciones == []


def test_parse_compra_con_air_codigo_332(tmp_path: Path):
    """Una compra con codRetAir=332 debe quedar accesible vía base_air_por_codigo."""
    xml = """<?xml version="1.0" encoding="UTF-8"?>
<iva>
    <IdInformante>1793205060001</IdInformante>
    <razonSocial>DEMO</razonSocial>
    <Anio>2026</Anio>
    <Mes>04</Mes>
    <compras>
        <detalleCompras>
            <codSustento>01</codSustento>
            <tpIdProv>02</tpIdProv>
            <idProv>0993222011001</idProv>
            <tipoComprobante>01</tipoComprobante>
            <baseNoGraIva>50.00</baseNoGraIva>
            <baseImponible>0.00</baseImponible>
            <baseImpGrav>1000.00</baseImpGrav>
            <montoIva>150.00</montoIva>
            <air>
                <detalleAir>
                    <codRetAir>332</codRetAir>
                    <baseImpAir>5856.01</baseImpAir>
                    <porcentajeAir>0.00</porcentajeAir>
                    <valRetAir>0.00</valRetAir>
                </detalleAir>
                <detalleAir>
                    <codRetAir>312</codRetAir>
                    <baseImpAir>1000.00</baseImpAir>
                    <porcentajeAir>1.00</porcentajeAir>
                    <valRetAir>10.00</valRetAir>
                </detalleAir>
            </air>
        </detalleCompras>
    </compras>
</iva>"""
    path = _write_xml(tmp_path, xml)
    ats = parse_ats(path)
    assert len(ats.compras) == 1
    compra = ats.compras[0]
    assert compra.id_prov == "0993222011001"
    assert compra.base_no_gra_iva == 50.00
    assert compra.base_imp_grav == 1000.00
    assert len(compra.air) == 2
    # Helpers
    assert ats.base_air_por_codigo("332") == 5856.01
    assert ats.valor_ret_por_codigo("312") == 10.00
    assert ats.total_base_no_gra_iva == 50.00
    assert ats.codigos_ir_presentes == {"332", "312"}


def test_parse_exportaciones_bienes_y_servicios(tmp_path: Path):
    """exportacionDe 01/02 → bienes (407), 03 → servicios (408)."""
    xml = """<?xml version="1.0" encoding="UTF-8"?>
<iva>
    <IdInformante>1793205060001</IdInformante>
    <Anio>2026</Anio>
    <Mes>04</Mes>
    <exportaciones>
        <detalleExportaciones>
            <exportacionDe>01</exportacionDe>
            <valorFOB>10000.00</valorFOB>
            <valorSeguro>500.00</valorSeguro>
            <valorFlete>200.00</valorFlete>
        </detalleExportaciones>
        <detalleExportaciones>
            <exportacionDe>02</exportacionDe>
            <valorFOB>5000.00</valorFOB>
            <valorSeguro>0.00</valorSeguro>
            <valorFlete>0.00</valorFlete>
        </detalleExportaciones>
        <detalleExportaciones>
            <exportacionDe>03</exportacionDe>
            <valorFOB>3000.00</valorFOB>
            <valorSeguro>0.00</valorSeguro>
            <valorFlete>0.00</valorFlete>
        </detalleExportaciones>
    </exportaciones>
</iva>"""
    path = _write_xml(tmp_path, xml)
    ats = parse_ats(path)
    assert ats.total_export_bienes == 10000 + 500 + 200 + 5000  # 01 + 02
    assert ats.total_export_servicios == 3000.00  # 03


def test_parse_xml_real_agroiris():
    """Sanity check con el ATS XML real (si existe localmente). Skip si no está."""
    real_path = Path(r"C:\Users\llper\OneDrive\Memin\Descargas\ATS-1793205060-2026-4.xml")
    if not real_path.exists():
        pytest.skip("ATS XML real no disponible en este entorno (esperado en CI).")
    ats = parse_ats(real_path)
    assert ats.ruc_informante == "1793205060001"
    assert len(ats.compras) == 453
    assert ats.base_air_por_codigo("332") == pytest.approx(10865.89, abs=0.01)
    assert ats.total_base_no_gra_iva == pytest.approx(220.11, abs=0.01)
