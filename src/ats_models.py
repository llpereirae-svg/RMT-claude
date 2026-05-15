"""Modelos para el Anexo Transaccional Simplificado (ATS) del SRI Ecuador.

Estructura per Ficha Tecnica ATS, root <iva>.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path


@dataclass
class ATSDetalleAir:
    """Detalle de retencion IR (Anexo Impuesto Renta) dentro de una compra."""
    cod_ret_air: str         # Codigo de concepto IR (303, 304, 312, 332, etc.)
    base_imp_air: float = 0.0
    porcentaje_air: float = 0.0
    val_ret_air: float = 0.0


@dataclass
class ATSCompra:
    cod_sustento: str = ""           # Tabla 5 (01-15, 00)
    tp_id_prov: str = ""             # 01=natural, 02=sociedad
    id_prov: str = ""                # RUC/CI del proveedor
    tipo_comprobante: str = ""       # Tabla 2 (01=factura, 02=nota venta, 15=cuotas, etc.)
    parte_rel: str = ""
    fecha_emision: str = ""          # dd/mm/yyyy
    establecimiento: str = ""
    punto_emision: str = ""
    secuencial: str = ""
    autorizacion: str = ""
    base_no_gra_iva: float = 0.0     # No grava IVA → casillero 531/541
    base_imponible: float = 0.0      # Base 0%
    base_imp_grav: float = 0.0       # Base gravada IVA
    base_imp_exe: float = 0.0        # Base exenta IR
    monto_ice: float = 0.0
    monto_iva: float = 0.0
    val_ret_bien10: float = 0.0      # Retencion IVA 10%
    val_ret_serv20: float = 0.0      # Retencion IVA 20%
    val_ret_serv50: float = 0.0      # Retencion IVA 50%
    val_ret_serv100: float = 0.0     # Retencion IVA 100%
    valor_ret_iva: float = 0.0
    valor_ret_ir: float = 0.0
    air: list[ATSDetalleAir] = field(default_factory=list)


@dataclass
class ATSVenta:
    tp_id_cliente: str = ""
    id_cliente: str = ""
    tipo_comprobante: str = ""
    tipo_emision: str = ""           # F=fisica, E=electronica
    numero_comprobantes: int = 0
    base_no_gra_iva: float = 0.0
    base_imponible: float = 0.0
    base_imp_grav: float = 0.0
    monto_iva: float = 0.0
    valor_ret_iva: float = 0.0
    valor_ret_ir: float = 0.0


@dataclass
class ATSExportacion:
    exportacion_de: str = ""         # 01,02=bienes ; 03=servicios u otros ingresos
    tip_ing_ext: str = ""            # Tipo de ingreso del exterior (tabla 18) si exportacion_de=03
    establecimiento: str = ""
    punto_emision: str = ""
    secuencial: str = ""
    autorizacion: str = ""
    fecha_emision: str = ""
    valor_fob: float = 0.0
    valor_seguro: float = 0.0
    valor_flete: float = 0.0

    @property
    def valor_total(self) -> float:
        return round(self.valor_fob + self.valor_seguro + self.valor_flete, 2)

    @property
    def es_bienes(self) -> bool:
        return self.exportacion_de in ("01", "02")

    @property
    def es_servicios(self) -> bool:
        return self.exportacion_de == "03"


@dataclass
class ATSAnulado:
    tipo_comprobante: str = ""
    establecimiento: str = ""
    punto_emision: str = ""
    secuencial: str = ""
    autorizacion: str = ""
    fecha_emision: str = ""


@dataclass
class ATSProcesado:
    archivo: Path
    nombre_archivo: str
    ruc_informante: str
    razon_social: str
    anio: str
    mes: str
    total_ventas: float = 0.0
    compras: list[ATSCompra] = field(default_factory=list)
    ventas: list[ATSVenta] = field(default_factory=list)
    exportaciones: list[ATSExportacion] = field(default_factory=list)
    anulados: list[ATSAnulado] = field(default_factory=list)

    # ─── Helpers para llenar casilleros del 104/103 ──────────────────────────
    @property
    def total_export_bienes(self) -> float:
        """Suma de exportaciones de bienes (exportacionDe 01 o 02) → casillero 407/417."""
        return round(sum(e.valor_total for e in self.exportaciones if e.es_bienes), 2)

    @property
    def total_export_servicios(self) -> float:
        """Suma de exportaciones de servicios (exportacionDe 03) → casillero 408/418."""
        return round(sum(e.valor_total for e in self.exportaciones if e.es_servicios), 2)

    @property
    def total_base_no_gra_iva(self) -> float:
        """Suma de baseNoGraIva en compras → casillero 531/541 exacto."""
        return round(sum(c.base_no_gra_iva for c in self.compras), 2)

    def base_air_por_codigo(self, codigo: str) -> float:
        """Suma de baseImpAir donde codRetAir = codigo. Util para Form 103 (ej. 332)."""
        return round(sum(a.base_imp_air for c in self.compras
                         for a in c.air if a.cod_ret_air == codigo), 2)

    def valor_ret_por_codigo(self, codigo: str) -> float:
        """Suma de valRetAir donde codRetAir = codigo. Casillero de retencion del 103."""
        return round(sum(a.val_ret_air for c in self.compras
                         for a in c.air if a.cod_ret_air == codigo), 2)

    @property
    def codigos_ir_presentes(self) -> set[str]:
        """Conjunto de codigos IR (codRetAir) que aparecen en al menos una compra."""
        return {a.cod_ret_air for c in self.compras for a in c.air if a.cod_ret_air}
