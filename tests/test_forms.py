"""Tests para el cálculo de casilleros 104 y 103."""
from pathlib import Path

from src.forms import calcular_formularios
from src.models import DocumentoRMT, RetencionRMT, RMTProcesado


def _rmt(documentos=None, retenciones=None) -> RMTProcesado:
    return RMTProcesado(
        archivo=Path("dummy.xlsx"),
        nombre_archivo="dummy.xlsx",
        hash_archivo="x",
        ruc="1793205060001",
        cliente_nombre="DEMO S.A.",
        periodo_inicio="2026-04",
        periodo_fin="2026-04",
        documentos=documentos or [],
        retenciones=retenciones or [],
    )


def _venta(base_0=0.0, base_iva=0.0, iva=0.0, sustento="") -> DocumentoRMT:
    return DocumentoRMT(
        bloque="FACTURAS EMITIDAS",
        grupo="venta",
        numero="001-001-001",
        fecha_emision=None,
        fecha_carga=None,
        contraparte_id="0912345678001",
        contraparte_nombre="CLIENTE",
        base_0=base_0,
        base_iva=base_iva,
        iva=iva,
        sustento=sustento,
    )


def _nc_venta(base_0=0.0, base_iva=0.0, iva=0.0) -> DocumentoRMT:
    return DocumentoRMT(
        bloque="NOTAS DE CREDITO EMITIDAS",
        grupo="nc_venta",
        numero="001-001-099",
        fecha_emision=None,
        fecha_carga=None,
        contraparte_id="0912345678001",
        contraparte_nombre="CLIENTE",
        base_0=base_0,
        base_iva=base_iva,
        iva=iva,
    )


def _compra(base_0=0.0, base_iva=0.0, iva=0.0) -> DocumentoRMT:
    return DocumentoRMT(
        bloque="DOCUMENTOS RECIBIDOS ATS",
        grupo="compra",
        numero="001-001-001",
        fecha_emision=None,
        fecha_carga=None,
        contraparte_id="0993222011001",
        contraparte_nombre="PROVEEDOR",
        base_0=base_0,
        base_iva=base_iva,
        iva=iva,
    )


# ─── Ventas ────────────────────────────────────────────────────────────────

def test_104_venta_tarifa_variable_simple():
    """Una factura con base IVA 1000 e IVA 150 (15%) → casilleros 410/420/430."""
    rmt = _rmt(documentos=[_venta(base_iva=1000.0, iva=150.0)])
    r = calcular_formularios(rmt)
    assert r.casilleros_104["410"] == 1000.0
    assert r.casilleros_104["420"] == 1000.0
    assert r.casilleros_104["430"] == 150.0


def test_104_venta_0_porciento_sin_credito_default():
    """Factura 0% (caso general) va a 403/413: tarifa 0% que NO da derecho a CT."""
    rmt = _rmt(documentos=[_venta(base_0=500.0)])
    r = calcular_formularios(rmt)
    assert r.casilleros_104["403"] == 500.0
    assert r.casilleros_104["413"] == 500.0
    # No debe escribir nada en 401/411 (tarifa ≠ 0)
    assert "401" not in r.casilleros_104
    assert "411" not in r.casilleros_104


def test_104_venta_0_porciento_con_credito_contri_especial():
    """Contribuyente especial: ventas 0% van a 405/415 (con derecho a CT)."""
    rmt = _rmt(documentos=[_venta(base_0=500.0)])
    r = calcular_formularios(rmt, ventas_0_con_credito=True)
    assert r.casilleros_104["405"] == 500.0
    assert r.casilleros_104["415"] == 500.0
    assert "403" not in r.casilleros_104


def test_104_nc_neteo_normal():
    """NC reduce la venta del mismo período."""
    rmt = _rmt(documentos=[
        _venta(base_iva=1000.0, iva=150.0),
        _nc_venta(base_iva=200.0, iva=30.0),
    ])
    r = calcular_formularios(rmt)
    assert r.casilleros_104["420"] == 800.0  # 1000 - 200
    assert r.casilleros_104["430"] == 120.0  # 150 - 30


def test_104_nc_mayor_que_venta_va_a_443():
    """Si NC supera las ventas, el exceso de base va al 443 (y el IVA al 453)."""
    rmt = _rmt(documentos=[
        _venta(base_iva=100.0, iva=15.0),
        _nc_venta(base_iva=300.0, iva=45.0),
    ])
    r = calcular_formularios(rmt)
    assert "420" not in r.casilleros_104  # net negativo, no se coloca
    assert r.casilleros_104["443"] == 200.0  # exceso de NC base
    assert r.casilleros_104["453"] == 30.0   # exceso de NC IVA


def test_104_exportacion_va_a_407():
    """Venta con sustento 'E' es exportación → casillero 407 (default bienes)."""
    rmt = _rmt(documentos=[_venta(base_0=5000.0, sustento="E")])
    r = calcular_formularios(rmt)
    assert r.casilleros_104["407"] == 5000.0
    assert r.casilleros_104["417"] == 5000.0


# ─── Compras ───────────────────────────────────────────────────────────────

def test_104_compra_con_iva_genera_credito():
    """Compra con base IVA 1000 e IVA 150 → 500/510/520 y crédito calculado."""
    rmt = _rmt(documentos=[
        _venta(base_iva=2000.0, iva=300.0),
        _compra(base_iva=1000.0, iva=150.0),
    ])
    r = calcular_formularios(rmt)
    assert r.casilleros_104["500"] == 1000.0
    assert r.casilleros_104["510"] == 1000.0
    assert r.casilleros_104["520"] == 150.0
    # Factor = 2000 / 2000 = 1.0, crédito = 150
    assert r.casilleros_104["563"] == 1.0
    assert r.casilleros_104["554"] == 150.0


def test_104_factor_proporcionalidad_solo_ventas_0_sin_credito():
    """Ventas 0% sin CT: factor 563 = 0, no hay crédito aplicable."""
    rmt = _rmt(documentos=[
        _venta(base_0=1000.0),
        _compra(base_iva=500.0, iva=75.0),
    ])
    r = calcular_formularios(rmt)
    assert r.casilleros_104.get("563", 0.0) == 0.0
    assert r.casilleros_104.get("554", 0.0) == 0.0
    # IVA de compras queda registrado en 520 aunque no se pueda usar como CT
    assert r.casilleros_104["520"] == 75.0


def test_104_factor_proporcionalidad_solo_ventas_0_con_credito():
    """Ventas 0% con CT (contri especial): factor 563 = 1.0, todo el IVA es crédito."""
    rmt = _rmt(documentos=[
        _venta(base_0=1000.0),
        _compra(base_iva=500.0, iva=75.0),
    ])
    r = calcular_formularios(rmt, ventas_0_con_credito=True)
    assert r.casilleros_104["563"] == 1.0
    assert r.casilleros_104["554"] == 75.0


def test_104_compras_0_porciento_van_a_507():
    """Compras locales 0% van al casillero 507/517 (no al 502/512)."""
    rmt = _rmt(documentos=[_compra(base_0=300.0)])
    r = calcular_formularios(rmt)
    assert r.casilleros_104["507"] == 300.0
    assert r.casilleros_104["517"] == 300.0
    assert "502" not in r.casilleros_104
    assert "512" not in r.casilleros_104


def test_104_importaciones_0_porciento_van_a_506():
    """Importaciones tarifa 0% van al casillero 506/516 (no al 505/515)."""
    importacion = DocumentoRMT(
        bloque="LIQUIDACION ADUANERA",
        grupo="importacion",
        numero="IMP-001",
        fecha_emision=None,
        fecha_carga=None,
        contraparte_id="000000000001",
        contraparte_nombre="ADUANA",
        base_0=2500.0,
    )
    rmt = _rmt(documentos=[importacion])
    r = calcular_formularios(rmt)
    assert r.casilleros_104["506"] == 2500.0
    assert r.casilleros_104["516"] == 2500.0
    assert "505" not in r.casilleros_104
    assert "515" not in r.casilleros_104


# ─── IVA retenido (agente de retención) ──────────────────────────────────

def test_104_retenciones_iva_emitidas_van_a_721_a_731():
    """Las retenciones emitidas se desagregan por porcentaje en casilleros 721-731."""
    rmt = _rmt(retenciones=[
        RetencionRMT(
            bloque="RETENCIONES EMITIDAS",
            numero="001-001-001",
            fecha_emision=None,
            contraparte_id="0993222011001",
            contraparte_nombre="PROV",
            iva_10=10.0, iva_30=30.0, iva_100=100.0,
        ),
    ])
    r = calcular_formularios(rmt)
    assert r.casilleros_104["721"] == 10.0
    assert r.casilleros_104["725"] == 30.0
    assert r.casilleros_104["731"] == 100.0
    assert r.casilleros_104["799"] == 140.0
    assert r.casilleros_104["801"] == 140.0


# ─── Formulario 103 ────────────────────────────────────────────────────────

def test_103_no_obligado_devuelve_vacio():
    rmt = _rmt(documentos=[_compra(base_iva=1000.0, iva=150.0)])
    r = calcular_formularios(rmt, obligado_103=False)
    assert r.casilleros_103 == {}


def test_103_obligado_con_empleados():
    rmt = _rmt()
    r = calcular_formularios(
        rmt,
        obligado_103=True,
        base_iess_empleados=2000.0,
        retencion_empleados=150.0,
    )
    assert r.casilleros_103["302"] == 2000.0
    assert r.casilleros_103["352"] == 150.0


def test_103_agrupa_retenciones_emitidas_por_codigo():
    """Dos retenciones con mismo código IR se suman en el mismo casillero."""
    rmt = _rmt(retenciones=[
        RetencionRMT(
            bloque="RETENCIONES EMITIDAS",
            numero="001",
            fecha_emision=None,
            contraparte_id="P1",
            contraparte_nombre="PROV",
            codigo_ir="312",
            base_ir=1000.0,
            valor_ir=10.0,
        ),
        RetencionRMT(
            bloque="RETENCIONES EMITIDAS",
            numero="002",
            fecha_emision=None,
            contraparte_id="P2",
            contraparte_nombre="PROV2",
            codigo_ir="312",
            base_ir=500.0,
            valor_ir=5.0,
        ),
    ])
    r = calcular_formularios(rmt, obligado_103=True)
    assert r.casilleros_103["312"] == 1500.0
    assert r.casilleros_103["362"] == 15.0  # 312 + 50
