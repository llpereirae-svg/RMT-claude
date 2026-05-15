"""Modelos internos para el procesador de RMT."""
from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path
from typing import Any


@dataclass
class DocumentoRMT:
    bloque: str
    grupo: str
    numero: str
    fecha_emision: date | None
    fecha_carga: datetime | None
    contraparte_id: str
    contraparte_nombre: str
    base_0: float = 0.0
    descuento_0: float = 0.0
    base_iva: float = 0.0
    descuento_iva: float = 0.0
    no_objeto_iva: float = 0.0
    descuento_no_objeto: float = 0.0
    total_sin_impuesto: float = 0.0
    iva: float = 0.0
    servicio: float = 0.0
    propina: float = 0.0
    total: float = 0.0
    sustento: str = ""
    autorizacion: str = ""
    doc_modificado: str = ""
    estado: str = "AUTORIZADO"
    cuadre: float = 0.0
    extra: dict[str, Any] = field(default_factory=dict)

    @property
    def base_0_neta(self) -> float:
        return round(self.base_0 - self.descuento_0, 2)

    @property
    def base_iva_neta(self) -> float:
        return round(self.base_iva - self.descuento_iva, 2)

    @property
    def no_objeto_neto(self) -> float:
        return round(self.no_objeto_iva - self.descuento_no_objeto + self.servicio, 2)


@dataclass
class RetencionRMT:
    bloque: str
    numero: str
    fecha_emision: date | None
    contraparte_id: str
    contraparte_nombre: str
    factura: str = ""
    iva_10: float = 0.0
    iva_20: float = 0.0
    iva_30: float = 0.0
    iva_50: float = 0.0
    iva_70: float = 0.0
    iva_100: float = 0.0
    codigo_ir: str = ""
    base_ir: float = 0.0
    valor_ir: float = 0.0
    extra: dict[str, Any] = field(default_factory=dict)


@dataclass
class AnuladoRMT:
    tipo: str
    fecha: date | None
    numero: str
    autorizacion: str
    contraparte_id: str
    contraparte_nombre: str
    total: float = 0.0


@dataclass
class RMTProcesado:
    archivo: Path
    nombre_archivo: str
    hash_archivo: str
    ruc: str
    cliente_nombre: str
    periodo_inicio: str
    periodo_fin: str
    documentos: list[DocumentoRMT] = field(default_factory=list)
    retenciones: list[RetencionRMT] = field(default_factory=list)
    anulados: list[AnuladoRMT] = field(default_factory=list)
    advertencias: list[str] = field(default_factory=list)

    @property
    def periodo_label(self) -> str:
        if self.periodo_inicio == self.periodo_fin:
            return self.periodo_inicio
        return f"{self.periodo_inicio} a {self.periodo_fin}"


@dataclass
class FormularioCalculado:
    casilleros_104: dict[str, float]
    casilleros_103: dict[str, float]
    resumen: dict[str, float]
    advertencias: list[str]
