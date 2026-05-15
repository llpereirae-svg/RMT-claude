"""Cálculo de casilleros 104/103 y exportación a Excel/PDF."""
from __future__ import annotations

import shutil
import subprocess
from collections import defaultdict
from pathlib import Path

import openpyxl

from .ats_models import ATSProcesado
from .models import DocumentoRMT, FormularioCalculado, RetencionRMT, RMTProcesado

# ─── Mapeo casillero → celda de valor en plantilla ───────────────────────────
# Patrón: label casillero está en columna X; el valor va en columna X+1.
# Generado inspeccionando FORMULARIO IVA.xlsx con openpyxl.
CELDA_104: dict[str, str] = {
    "401": "J9",  "411": "L9",  "421": "N9",
    "402": "J10", "412": "L10", "422": "N10",
    "410": "J11", "420": "L11", "430": "N11",
    "425": "J12", "435": "L12", "445": "N12",
    "423": "N13", "424": "N14",
    "403": "J15", "413": "L15",
    "404": "J16", "414": "L16",
    "405": "J17", "415": "L17",
    "406": "J18", "416": "L18",
    "407": "J19", "417": "L19",
    "408": "J20", "418": "L20",
    "409": "J21", "419": "L21", "429": "N21",
    "431": "J22", "441": "L22",
    "442": "L23",
    "443": "L24", "453": "N24",
    "434": "J25", "444": "L25", "454": "N25",
    "480": "N28", "481": "N29", "482": "N30",
    "483": "N31", "484": "N32", "485": "N33",
    "486": "N34", "487": "N35", "499": "N36",
    "111": "F38", "113": "N38",
    "500": "J42", "510": "L42", "520": "N42",
    "501": "J43", "511": "L43", "521": "N43",
    "530": "J44", "533": "L44", "534": "N44",
    "540": "J45", "550": "L45", "560": "N45",
    "502": "J46", "512": "L46", "522": "N46",
    "503": "J47", "513": "L47", "523": "N47",
    "504": "J48", "514": "L48", "524": "N48",
    "505": "J49", "515": "L49", "525": "N49",
    "526": "N50", "527": "N51",
    "506": "J52", "516": "L52",
    "507": "J53", "517": "L53",
    "508": "J54", "518": "L54",
    "509": "J55", "519": "L55", "529": "N55",
    "531": "J56", "541": "L56",
    "532": "J57", "542": "L57",
    "543": "L58",
    "544": "L59", "554": "N59",
    "535": "J60", "545": "L60", "555": "N60",
    "563": "N62", "564": "N63", "565": "N64",
    "115": "F66", "117": "N66", "119": "N67",
    "601": "N70", "602": "N71", "603": "N72",
    "604": "N73", "605": "N75", "606": "N76",
    "607": "N77", "608": "N78", "623": "N79",
    "609": "N81", "622": "N82",
    "615": "N89", "617": "N90", "618": "N91",
    "619": "N92", "620": "N96", "699": "N98",
    "721": "N107", "723": "N108", "725": "N109",
    "727": "N110", "729": "N111", "731": "N112",
    "799": "N113", "800": "N114", "802": "N115",
    "801": "N116",
    "859": "N118",
    "902": "N137",
}

CELDA_103: dict[str, str] = {
    "302": "L5",  "352": "N5",
    "303": "L7",  "353": "N7",
    "304": "L9",  "354": "N9",
    "307": "L10", "357": "N10",
    "308": "L11", "358": "N11",
    "309": "L12", "359": "N12",
    "310": "L13", "360": "N13",
    "311": "L14", "361": "N14",
    "312": "L16", "362": "N16",
    "322": "L17", "372": "N17",
    "343": "L21", "393": "N21",
    "344": "L22", "394": "N22",
    "332": "L23",
    "314": "L25", "364": "N25",
    "319": "L28", "369": "N28",
    "320": "L29", "370": "N29",
    "323": "L31", "373": "N31",
    "325": "L35", "375": "N35",
    "326": "L37", "376": "N37",
    "327": "L38", "377": "N38",
    "328": "L39", "378": "N39",
    "335": "L46", "385": "N46",
}


def calcular_formularios(
    rmt: RMTProcesado,
    *,
    ats: ATSProcesado | None = None,
    obligado_103: bool = False,
    base_iess_empleados: float = 0.0,
    retencion_empleados: float = 0.0,
    ventas_0_con_credito: bool = False,
) -> FormularioCalculado:
    """Calcula casilleros 104 y 103.

    Args:
        ats: si se provee, se usa para refinar 407/408 (exportaciones bienes vs servicios),
            531/541 (baseNoGraIva exacta) y 332 del F103 (codRetAir).
        ventas_0_con_credito: si True, las ventas locales 0% van al casillero 405/415
            (contribuyente especial o con actividad que da derecho a CT). Si False (default),
            van al 403/413 (sin derecho a CT, caso general).
    """
    cas104 = _calcular_104(rmt, ats=ats, ventas_0_con_credito=ventas_0_con_credito)
    cas103 = _calcular_103(rmt, obligado_103, base_iess_empleados, retencion_empleados, ats=ats)
    resumen = {
        "total_ventas_neto": cas104.get("419", 0.0),
        "iva_generado": cas104.get("430", 0.0),
        "total_compras_neto": cas104.get("519", 0.0),
        "iva_credito": cas104.get("554", 0.0),
        "iva_a_pagar": cas104.get("564", 0.0),
        "iva_a_favor": cas104.get("565", 0.0),
        "total_retenido_iva": cas104.get("799", 0.0),
        "total_104": cas104.get("859", 0.0),
        "total_103": cas103.get("699", 0.0),
    }
    advertencias = list(rmt.advertencias)
    return FormularioCalculado(
        casilleros_104=cas104,
        casilleros_103=cas103,
        resumen=resumen,
        advertencias=advertencias,
    )


# ─── Cálculo Formulario 104 ────────────────────────────────────────────────────

def _is_nota_venta(d: DocumentoRMT) -> bool:
    """Una Nota de Venta indica que el proveedor es RISE / Negocio Popular.
    En el RMT aparece en la columna TIPO (DOC ATS) o DOC (GASTOS DE VIAJE)."""
    tipo = (d.extra.get("TIPO") or d.extra.get("DOC") or "").upper()
    return "NOTA DE VENTA" in tipo


def _is_gasto_deducible_sin_factura(d: DocumentoRMT) -> bool:
    """Tipo de comprobante 'GASTO DEDUCIBLE SIN FACTURA' en gastos de viaje:
    no se considera para el cálculo del 104 (per regla del contador)."""
    doc = (d.extra.get("DOC") or "").upper()
    return "GASTO DEDUCIBLE SIN FACTURA" in doc


# Bloques que NO se consideran para el cálculo del 104 (per regla del contador).
BLOQUES_EXCLUIDOS_104: frozenset[str] = frozenset({
    "DOCUMENTOS RECIBIDOS QUE NO VAN EN EL ATS",
    "FACTURAS RECIBIDAS (CEDULA)",
})


def _calcular_104(
    rmt: RMTProcesado,
    *,
    ats: ATSProcesado | None = None,
    ventas_0_con_credito: bool = False,
) -> dict[str, float]:
    c: dict[str, float] = {}

    ventas = [d for d in rmt.documentos if d.grupo == "venta"]
    nc_ventas = [d for d in rmt.documentos if d.grupo == "nc_venta"]
    # Filtrar compras: excluir bloques no declarables, gastos deducibles sin factura,
    # y separar Notas de Venta (RISE/NP).
    compras_todas = [d for d in rmt.documentos
                     if d.grupo == "compra"
                     and d.bloque not in BLOQUES_EXCLUIDOS_104
                     and not _is_gasto_deducible_sin_factura(d)]
    rise_np = [d for d in compras_todas if _is_nota_venta(d)]
    compras = [d for d in compras_todas if not _is_nota_venta(d)]
    nc_compras = [d for d in rmt.documentos if d.grupo == "nc_compra"]
    importaciones = [d for d in rmt.documentos if d.grupo == "importacion"]

    ret_emit = [r for r in rmt.retenciones if "EMITIDAS" in r.bloque]

    # ── VENTAS ──────────────────────────────────────────────────────────────
    exports = [d for d in ventas if d.sustento.strip().upper() == "E"]
    loc_ventas = [d for d in ventas if d.sustento.strip().upper() != "E"]

    # Casilleros 403/413 vs 405/415 segun el regimen del contribuyente.
    # 403/413: tarifa 0% que NO da derecho a credito tributario (default).
    # 405/415: tarifa 0% que SI da derecho a CT (contribuyente especial / actividad especial).
    cas_0_gross = "405" if ventas_0_con_credito else "403"
    cas_0_net = "415" if ventas_0_con_credito else "413"

    # Ventas locales 0% (gross y neto)
    gross_0 = _s(loc_ventas, "base_0") - _s(loc_ventas, "descuento_0")
    nc_0 = _s(nc_ventas, "base_0") - _s(nc_ventas, "descuento_0")
    net_0 = round(gross_0 - nc_0, 2)
    if net_0 >= 0:
        c[cas_0_gross] = gross_0
        c[cas_0_net] = net_0
    else:
        c["442"] = round(abs(net_0), 2)

    # Ventas locales tarifa variable (15%, RMT la rotula 12% por legacy)
    gross_iva = _s(loc_ventas, "base_iva") - _s(loc_ventas, "descuento_iva")
    nc_iva_base = _s(nc_ventas, "base_iva") - _s(nc_ventas, "descuento_iva")
    net_iva = round(gross_iva - nc_iva_base, 2)
    iva_emit = _s(loc_ventas, "iva")
    nc_iva_monto = _s(nc_ventas, "iva")
    net_iva_monto = round(iva_emit - nc_iva_monto, 2)
    if net_iva >= 0:
        c["410"] = gross_iva
        c["420"] = net_iva
        c["430"] = max(0.0, net_iva_monto)
    else:
        c["443"] = round(abs(net_iva), 2)
        if nc_iva_monto > iva_emit:
            c["453"] = round(nc_iva_monto - iva_emit, 2)

    # No objeto IVA (incluye servicio y propina per instrucción usuario)
    no_obj = _s(loc_ventas, "no_objeto_iva") - _s(loc_ventas, "descuento_no_objeto") + _s(loc_ventas, "servicio") + _s(loc_ventas, "propina")
    nc_no_obj = _s(nc_ventas, "no_objeto_iva") - _s(nc_ventas, "descuento_no_objeto")
    net_no_obj = round(no_obj - nc_no_obj, 2)
    if net_no_obj > 0:
        c["431"] = net_no_obj

    # Exportaciones → si hay ATS, separar bienes (407) vs servicios (408) por exportacionDe.
    # Sin ATS, asumir bienes (caso comun) per fallback.
    export_total = _s(exports, "base_0") + _s(exports, "base_iva")
    if ats is not None and (ats.total_export_bienes + ats.total_export_servicios) > 0:
        if ats.total_export_bienes > 0:
            c["407"] = ats.total_export_bienes
            c["417"] = ats.total_export_bienes
        if ats.total_export_servicios > 0:
            c["408"] = ats.total_export_servicios
            c["418"] = ats.total_export_servicios
    elif export_total > 0:
        c["407"] = round(export_total, 2)
        c["417"] = c["407"]

    # Si hay ATS, el total de exportaciones es el oficial (FOB+seguro+flete).
    export_total_eff = (
        ats.total_export_bienes + ats.total_export_servicios
        if ats is not None and (ats.total_export_bienes + ats.total_export_servicios) > 0
        else export_total
    )

    # Totales ventas: suma de tarifa variable + 0% local + exportaciones (no_objeto va aparte en 431)
    c["409"] = round(gross_iva + gross_0 + export_total_eff, 2)
    c["419"] = round(max(0, net_iva) + max(0, net_0) + export_total_eff, 2)
    if c.get("430"):
        c["429"] = c["430"]

    # ── COMPRAS ─────────────────────────────────────────────────────────────
    # Local con IVA (asume crédito tributario total — refinable con ATS)
    g_iva_c = _s(compras, "base_iva") - _s(compras, "descuento_iva")
    nc_iva_c = _s(nc_compras, "base_iva") - _s(nc_compras, "descuento_iva")
    net_iva_c = round(g_iva_c - nc_iva_c, 2)
    iva_c = round(_s(compras, "iva") - _s(nc_compras, "iva"), 2)
    c["500"] = round(g_iva_c, 2)
    # Si la NC ya quedó neteada en 510, no se declara 544. Solo si excede al gross
    # (caso analogo al patron 442/443 en ventas).
    if net_iva_c >= 0:
        c["510"] = net_iva_c
        c["520"] = max(0.0, iva_c)
    else:
        c["510"] = 0.0
        c["544"] = round(abs(net_iva_c), 2)
        c["520"] = 0.0
        c["554"] = 0.0  # se ajusta luego en liquidación

    # Adquisiciones locales tarifa 0% → casillero 507/517 (incluye AF tarifa 0%)
    g_0_c = _s(compras, "base_0") - _s(compras, "descuento_0")
    nc_0_c = _s(nc_compras, "base_0") - _s(nc_compras, "descuento_0")
    net_0_c = round(g_0_c - nc_0_c, 2)
    if g_0_c > 0:
        c["507"] = round(g_0_c, 2)
        c["517"] = net_0_c
    if nc_0_c > 0:
        c["543"] = round(nc_0_c, 2)

    # Adquisiciones a RISE / Negocios Populares → casillero 508/518
    g_rise = _s(rise_np, "base_0") - _s(rise_np, "descuento_0")
    if g_rise > 0:
        c["508"] = round(g_rise, 2)
        c["518"] = round(g_rise, 2)

    # No objeto IVA en compras → casillero 531 (no entra en el total 509).
    # Si hay ATS, usar suma exacta de baseNoGraIva en <compras> (es el dato declarado oficialmente).
    # Sin ATS, aproximar desde el RMT: NO OBJETO DE IVA + propina + % servicio.
    if ats is not None and ats.total_base_no_gra_iva > 0:
        no_obj_c = ats.total_base_no_gra_iva
    else:
        todas_para_no_obj = compras + rise_np
        no_obj_c = round(
            _s(todas_para_no_obj, "no_objeto_iva") - _s(todas_para_no_obj, "descuento_no_objeto")
            + _s(todas_para_no_obj, "propina")
            + _s(todas_para_no_obj, "servicio"),
            2,
        )
    if no_obj_c > 0:
        c["531"] = no_obj_c
        c["541"] = no_obj_c

    # Importaciones (Liquidación Aduanera)
    imp_base_iva = round(_s(importaciones, "base_iva"), 2)
    imp_base_0 = round(_s(importaciones, "base_0"), 2)
    imp_iva = round(_s(importaciones, "iva"), 2)
    if imp_base_iva > 0:
        # Importaciones de bienes (excluye AF) tarifa ≠ 0
        c["504"] = imp_base_iva
        c["514"] = imp_base_iva
        c["524"] = imp_iva
    if imp_base_0 > 0:
        # Importaciones de bienes (incluye AF) tarifa 0%
        c["506"] = imp_base_0
        c["516"] = imp_base_0

    # Total adquisiciones gravadas: 500 + 504 + 506 + 507 + 508 (no incluye 531/532)
    c["509"] = round(g_iva_c + g_0_c + g_rise + imp_base_iva + imp_base_0, 2)
    c["519"] = round(net_iva_c + max(0, net_0_c) + g_rise + imp_base_iva + imp_base_0, 2)
    c["529"] = round(c.get("520", 0) + imp_iva, 2)

    # ── LIQUIDACIÓN ──────────────────────────────────────────────────────────
    total_ventas = c.get("419", 0.0)
    # Ventas con derecho a CT = tarifa variable + exportaciones (+ ventas 0% si aplica)
    ventas_con_credito = max(0, net_iva) + export_total_eff
    if ventas_0_con_credito:
        ventas_con_credito += max(0, net_0)
    factor = round(ventas_con_credito / total_ventas, 4) if total_ventas > 0 else 0.0
    c["563"] = factor

    iva_credito = round((c.get("520", 0) + c.get("524", 0)) * factor, 2)
    c["554"] = iva_credito

    iva_neto = round(c.get("430", 0) - iva_credito, 2)
    if iva_neto > 0:
        c["564"] = iva_neto
        c["601"] = iva_neto
    else:
        c["565"] = round(abs(iva_neto), 2)
        c["602"] = c["565"]

    # ── IVA RETENCIONES EMITIDAS (agente de retención) → Sección 720-799 ───
    c["721"] = round(sum(r.iva_10 for r in ret_emit), 2)
    c["723"] = round(sum(r.iva_20 for r in ret_emit), 2)
    c["725"] = round(sum(r.iva_30 for r in ret_emit), 2)
    c["727"] = round(sum(r.iva_50 for r in ret_emit), 2)
    c["729"] = round(sum(r.iva_70 for r in ret_emit), 2)
    c["731"] = round(sum(r.iva_100 for r in ret_emit), 2)
    c["799"] = round(sum(c.get(k, 0) for k in ["721", "723", "725", "727", "729", "731"]), 2)
    c["801"] = c["799"]

    # Total consolidado 859 = 699 (percepción, generalmente = 601) + 801
    c["699"] = c.get("601", 0.0)
    c["859"] = round(c.get("699", 0) + c.get("801", 0), 2)
    # Total a pagar final 902 (= 859 cuando no hay pagos previos en 898)
    c["902"] = c["859"]

    # ── CONTEO DE COMPROBANTES ───────────────────────────────────────────────
    emit_ok = [d for d in rmt.documentos if d.grupo in ("venta",) and d.estado != "ANULADO"]
    anulados_emit = [a for a in rmt.anulados if "EMITIDAS" in a.tipo or "EMITIDOS" in a.tipo]
    nc_recib = [d for d in rmt.documentos if d.grupo == "nc_compra"]
    liq_compra = [d for d in compras if "LIQUIDACION DE COMPRA" in d.bloque]
    c["115"] = float(len(emit_ok))
    c["113"] = float(len(anulados_emit))
    c["117"] = float(len(nc_recib))
    c["119"] = float(len(liq_compra))

    return {k: round(v, 2) for k, v in c.items() if v != 0.0}


# ─── Cálculo Formulario 103 ────────────────────────────────────────────────────

def _calcular_103(
    rmt: RMTProcesado,
    obligado: bool,
    base_iess: float,
    ret_empleados: float,
    *,
    ats: ATSProcesado | None = None,
) -> dict[str, float]:
    if not obligado:
        return {}

    c: dict[str, float] = {}
    ret_emit = [r for r in rmt.retenciones if "EMITIDAS" in r.bloque]
    compras = [d for d in rmt.documentos if d.grupo in ("compra", "importacion")]
    gastos_viaje = [d for d in rmt.documentos if d.bloque == "GASTOS DE VIAJE"]

    # Empleados (relación de dependencia)
    if base_iess > 0:
        c["302"] = round(base_iess, 2)
        c["352"] = round(ret_empleados, 2)

    # ── Si hay ATS: usar codRetAir como fuente oficial de bases/valores por código IR ──
    if ats is not None and ats.codigos_ir_presentes:
        for cod in ats.codigos_ir_presentes:
            base = ats.base_air_por_codigo(cod)
            valor = ats.valor_ret_por_codigo(cod)
            if base > 0 and cod in CELDA_103:
                c[cod] = base
            if valor > 0:
                try:
                    ret_cas = str(int(cod) + 50)
                    if ret_cas in CELDA_103:
                        c[ret_cas] = valor
                except ValueError:
                    pass
    else:
        # Fallback sin ATS: agrupar retenciones emitidas del RMT por código IR
        por_codigo: dict[str, dict[str, float]] = defaultdict(lambda: {"base": 0.0, "valor": 0.0})
        for r in ret_emit:
            if r.codigo_ir:
                por_codigo[r.codigo_ir]["base"] += r.base_ir
                por_codigo[r.codigo_ir]["valor"] += r.valor_ir
        for cod, montos in por_codigo.items():
            if montos["base"] > 0 and cod in CELDA_103:
                c[cod] = round(montos["base"], 2)
            if montos["valor"] > 0:
                try:
                    ret_cas = str(int(cod) + 50)
                    if ret_cas in CELDA_103:
                        c[ret_cas] = round(montos["valor"], 2)
                except ValueError:
                    pass

        # Casillero 332 fallback: total_retenido + compras_netas - RG - propinas en compras
        total_ret_valor = sum(r.valor_ir for r in ret_emit)
        compras_netas_base = sum(d.base_0 + d.base_iva - d.descuento_0 - d.descuento_iva for d in compras)
        total_rg = sum(d.total for d in gastos_viaje)
        propinas_compras = sum(d.propina for d in compras)
        base_332 = round(total_ret_valor + compras_netas_base - total_rg - propinas_compras, 2)
        if base_332 > 0:
            c["332"] = base_332

    # Totales Form 103 (solo casilleros con clave numérica pura)
    def _intval(k: str) -> int | None:
        try:
            return int(k)
        except ValueError:
            return None

    base_cas = {k for k in c if (_intval(k) is not None) and _intval(k) < 400}
    ret_cas_set = {k for k in c if (_intval(k) is not None) and 350 <= _intval(k) < 500}
    c["499"] = round(sum(c[k] for k in base_cas), 2)
    c["549"] = round(sum(c[k] for k in ret_cas_set if k not in base_cas), 2)
    c["699"] = c.get("549", 0.0)

    return {k: round(v, 2) for k, v in c.items() if v != 0.0}


# ─── Exportación a Excel y PDF ────────────────────────────────────────────────

def exportar_formulario(
    template_path: Path,
    output_path: Path,
    formularios: FormularioCalculado,
) -> Path:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(template_path, output_path)

    wb = openpyxl.load_workbook(output_path)

    _fill_sheet(wb["104"], formularios.casilleros_104, CELDA_104)
    if formularios.casilleros_103:
        _fill_sheet(wb["103"], formularios.casilleros_103, CELDA_103)

    wb.save(output_path)

    _try_export_pdf(output_path)
    return output_path


def _fill_sheet(ws, casilleros: dict[str, float], mapa: dict[str, str]) -> None:
    for cas, valor in casilleros.items():
        celda = mapa.get(str(cas))
        if celda and valor != 0.0:
            ws[celda] = round(valor, 2)


def _try_export_pdf(xlsx_path: Path) -> Path | None:
    pdf_path = xlsx_path.with_suffix(".pdf")
    # Intento 1: LibreOffice headless
    lo = shutil.which("soffice") or shutil.which("libreoffice")
    if lo:
        try:
            subprocess.run(
                [lo, "--headless", "--convert-to", "pdf", "--outdir", str(xlsx_path.parent), str(xlsx_path)],
                timeout=30,
                capture_output=True,
            )
            if pdf_path.exists():
                return pdf_path
        except Exception:
            pass
    # Intento 2: Excel via COM (Windows)
    try:
        import win32com.client as win32  # type: ignore
        xl = win32.Dispatch("Excel.Application")
        xl.Visible = False
        wb = xl.Workbooks.Open(str(xlsx_path.resolve()))
        wb.ExportAsFixedFormat(0, str(pdf_path.resolve()))
        wb.Close(False)
        xl.Quit()
        return pdf_path
    except Exception:
        pass
    return None


# ─── Helpers ──────────────────────────────────────────────────────────────────

def _s(items: list, attr: str) -> float:
    return round(sum(getattr(d, attr, 0.0) or 0.0 for d in items), 2)
