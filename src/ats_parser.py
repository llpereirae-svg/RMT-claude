"""Parser del Anexo Transaccional Simplificado (ATS) en formato XML.

Estructura per Ficha Tecnica ATS del SRI Ecuador (root <iva>).
"""
from __future__ import annotations

from pathlib import Path

from lxml import etree

from .ats_models import (
    ATSAnulado,
    ATSCompra,
    ATSDetalleAir,
    ATSExportacion,
    ATSProcesado,
    ATSVenta,
)


def _num(value: str | None) -> float:
    if value is None or value == "":
        return 0.0
    try:
        return float(value)
    except (ValueError, TypeError):
        return 0.0


def _txt(value: str | None) -> str:
    return (value or "").strip()


def parse_ats(path: Path) -> ATSProcesado:
    """Parsea un ATS XML del SRI Ecuador.

    Acepta tanto root <iva> (formato actual) como <ats> (formato antiguo).
    """
    tree = etree.parse(str(path))
    root = tree.getroot()

    parsed = ATSProcesado(
        archivo=path,
        nombre_archivo=path.name,
        ruc_informante=_txt(root.findtext("IdInformante")),
        razon_social=_txt(root.findtext("razonSocial")),
        anio=_txt(root.findtext("Anio")),
        mes=_txt(root.findtext("Mes")),
        total_ventas=_num(root.findtext("totalVentas")),
    )

    # Compras (cada una puede tener un sub-bloque <air><detalleAir/></air>)
    for c in root.iterfind("compras/detalleCompras"):
        air_list: list[ATSDetalleAir] = []
        for a in c.iterfind("air/detalleAir"):
            air_list.append(ATSDetalleAir(
                cod_ret_air=_txt(a.findtext("codRetAir")),
                base_imp_air=_num(a.findtext("baseImpAir")),
                porcentaje_air=_num(a.findtext("porcentajeAir")),
                val_ret_air=_num(a.findtext("valRetAir")),
            ))
        parsed.compras.append(ATSCompra(
            cod_sustento=_txt(c.findtext("codSustento")),
            tp_id_prov=_txt(c.findtext("tpIdProv")),
            id_prov=_txt(c.findtext("idProv")),
            tipo_comprobante=_txt(c.findtext("tipoComprobante")),
            parte_rel=_txt(c.findtext("parteRel")),
            fecha_emision=_txt(c.findtext("fechaEmision")),
            establecimiento=_txt(c.findtext("establecimiento")),
            punto_emision=_txt(c.findtext("puntoEmision")),
            secuencial=_txt(c.findtext("secuencial")),
            autorizacion=_txt(c.findtext("autorizacion")),
            base_no_gra_iva=_num(c.findtext("baseNoGraIva")),
            base_imponible=_num(c.findtext("baseImponible")),
            base_imp_grav=_num(c.findtext("baseImpGrav")),
            base_imp_exe=_num(c.findtext("baseImpExe")),
            monto_ice=_num(c.findtext("montoIce")),
            monto_iva=_num(c.findtext("montoIva")),
            val_ret_bien10=_num(c.findtext("valRetBien10")),
            val_ret_serv20=_num(c.findtext("valRetServ20")),
            val_ret_serv50=_num(c.findtext("valRetServ50")),
            val_ret_serv100=_num(c.findtext("valRetServ100")),
            valor_ret_iva=_num(c.findtext("valorRetIva")),
            valor_ret_ir=_num(c.findtext("valorRetIr")),
            air=air_list,
        ))

    # Ventas
    for v in root.iterfind("ventas/detalleVentas"):
        parsed.ventas.append(ATSVenta(
            tp_id_cliente=_txt(v.findtext("tpIdCliente")),
            id_cliente=_txt(v.findtext("idCliente")),
            tipo_comprobante=_txt(v.findtext("tipoComprobante")),
            tipo_emision=_txt(v.findtext("tipoEmision")),
            numero_comprobantes=int(_num(v.findtext("numeroComprobantes"))),
            base_no_gra_iva=_num(v.findtext("baseNoGraIva")),
            base_imponible=_num(v.findtext("baseImponible")),
            base_imp_grav=_num(v.findtext("baseImpGrav")),
            monto_iva=_num(v.findtext("montoIva")),
            valor_ret_iva=_num(v.findtext("valorRetIva")),
            valor_ret_ir=_num(v.findtext("valorRetIr")),
        ))

    # Exportaciones
    for e in root.iterfind("exportaciones/detalleExportaciones"):
        parsed.exportaciones.append(ATSExportacion(
            exportacion_de=_txt(e.findtext("exportacionDe")),
            tip_ing_ext=_txt(e.findtext("tipIngExt")),
            establecimiento=_txt(e.findtext("establecimiento")),
            punto_emision=_txt(e.findtext("puntoEmision")),
            secuencial=_txt(e.findtext("secuencial")),
            autorizacion=_txt(e.findtext("autorizacion")),
            fecha_emision=_txt(e.findtext("fechaEmision")),
            valor_fob=_num(e.findtext("valorFOB")),
            valor_seguro=_num(e.findtext("valorSeguro")),
            valor_flete=_num(e.findtext("valorFlete")),
        ))

    # Anulados
    for a in root.iterfind("anulados/detalleAnulados"):
        parsed.anulados.append(ATSAnulado(
            tipo_comprobante=_txt(a.findtext("tipoComprobante")),
            establecimiento=_txt(a.findtext("establecimiento")),
            punto_emision=_txt(a.findtext("puntoEmision")),
            secuencial=_txt(a.findtext("secuencial")),
            autorizacion=_txt(a.findtext("autorizacion")),
            fecha_emision=_txt(a.findtext("fechaEmision")),
        ))

    return parsed
