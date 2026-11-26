from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date
from decimal import Decimal, ROUND_HALF_UP
from typing import Literal, Optional
from zoneinfo import ZoneInfo

ZONA_ARG = ZoneInfo("America/Argentina/Buenos_Aires")
MODO_REDONDEO = ROUND_HALF_UP
DOS_DECIMALES = Decimal("0.01")


@dataclass(frozen=True)
class TipoComprobanteARCA:
    codigo: int
    descripcion: str
    letra: str
    signo: int  # +1 suma (Factura / ND / Recibo), -1 resta (Nota de Crédito)
    discrimina_iva: bool
    es_factura_m: bool = False
    es_comprobante_fiscal: bool = True


TIPOS_COMPROBANTE_ARCA: dict[int, TipoComprobanteARCA] = {
    1: TipoComprobanteARCA(1, "FACTURA A", "A", 1, True),
    2: TipoComprobanteARCA(2, "NOTA DE DEBITO A", "A", 1, True),
    3: TipoComprobanteARCA(3, "NOTA DE CREDITO A", "A", -1, True),
    4: TipoComprobanteARCA(4, "RECIBO A", "A", 1, True),
    6: TipoComprobanteARCA(6, "FACTURA B", "B", 1, False),
    7: TipoComprobanteARCA(7, "NOTA DE DEBITO B", "B", 1, False),
    8: TipoComprobanteARCA(8, "NOTA DE CREDITO B", "B", -1, False),
    9: TipoComprobanteARCA(9, "RECIBO B", "B", 1, False),
    11: TipoComprobanteARCA(11, "FACTURA C", "C", 1, False),
    12: TipoComprobanteARCA(12, "NOTA DE DEBITO C", "C", 1, False),
    13: TipoComprobanteARCA(13, "NOTA DE CREDITO C", "C", -1, False),
    15: TipoComprobanteARCA(15, "RECIBO C", "C", 1, False),
    51: TipoComprobanteARCA(51, "FACTURA M", "M", 1, True, es_factura_m=True),
    52: TipoComprobanteARCA(52, "NOTA DE DEBITO M", "M", 1, True, es_factura_m=True),
    53: TipoComprobanteARCA(53, "NOTA DE CREDITO M", "M", -1, True, es_factura_m=True),
    54: TipoComprobanteARCA(54, "RECIBO M", "M", 1, True, es_factura_m=True),
    91: TipoComprobanteARCA(91, "REMITO R", "R", 0, False, es_comprobante_fiscal=False),
    201: TipoComprobanteARCA(201, "FACTURA DE CREDITO ELECTRONICA MiPyME A", "A", 1, True),
    202: TipoComprobanteARCA(202, "NOTA DE DEBITO ELECTRONICA MiPyME A", "A", 1, True),
    203: TipoComprobanteARCA(203, "NOTA DE CREDITO ELECTRONICA MiPyME A", "A", -1, True),
    206: TipoComprobanteARCA(206, "FACTURA DE CREDITO ELECTRONICA MiPyME B", "B", 1, False),
    207: TipoComprobanteARCA(207, "NOTA DE DEBITO ELECTRONICA MiPyME B", "B", 1, False),
    208: TipoComprobanteARCA(208, "NOTA DE CREDITO ELECTRONICA MiPyME B", "B", -1, False),
    211: TipoComprobanteARCA(211, "FACTURA DE CREDITO ELECTRONICA MiPyME C", "C", 1, False),
    212: TipoComprobanteARCA(212, "NOTA DE DEBITO ELECTRONICA MiPyME C", "C", 1, False),
    213: TipoComprobanteARCA(213, "NOTA DE CREDITO ELECTRONICA MiPyME C", "C", -1, False),
}


@dataclass(frozen=True)
class AlicuotaIVAConfig:
    codigo_arca: int
    tasa: Decimal
    descripcion: str
    vigente_desde: date
    vigente_hasta: Optional[date] = None

    def esta_vigente(self, fecha: date) -> bool:
        if fecha < self.vigente_desde:
            return False
        if self.vigente_hasta is not None and fecha > self.vigente_hasta:
            return False
        return True


ALICUOTAS_IVA_DEFAULT: list[AlicuotaIVAConfig] = [
    AlicuotaIVAConfig(3, Decimal("0.00"), "IVA 0%", date(2000, 1, 1)),
    AlicuotaIVAConfig(9, Decimal("2.50"), "IVA 2,5%", date(2014, 1, 1)),
    AlicuotaIVAConfig(8, Decimal("5.00"), "IVA 5%", date(2014, 1, 1)),
    AlicuotaIVAConfig(4, Decimal("10.50"), "IVA 10,5%", date(2000, 1, 1)),
    AlicuotaIVAConfig(5, Decimal("21.00"), "IVA 21%", date(2000, 1, 1)),
    AlicuotaIVAConfig(6, Decimal("27.00"), "IVA 27%", date(2000, 1, 1)),
]


@dataclass(frozen=True)
class RegimenRetencion:
    nombre: str
    impuesto: Literal["GANANCIAS", "IVA", "IIBB", "SUSS"]
    jurisdiccion: str
    alicuota: Decimal
    minimo_no_sujeto: Decimal
    base_calculo: Literal["NETO_GRAVADO", "TOTAL", "IVA"]
    vigente_desde: date
    vigente_hasta: Optional[date] = None
    usa_padron: bool = False

    def esta_vigente(self, fecha: date) -> bool:
        if fecha < self.vigente_desde:
            return False
        if self.vigente_hasta is not None and fecha > self.vigente_hasta:
            return False
        return True


@dataclass
class ParametrosFiscales:
    cuit_empresa: str = "30716137488"  # 3CICON SAS
    razon_social_empresa: str = "3CICON SAS"
    jurisdiccion_principal: str = "PBA"
    moneda_operativa: str = "PES"
    tolerancia_reconciliacion: Decimal = Decimal("0.06")
    tolerancia_iva_alicuota: Decimal = Decimal("0.06")
    politica_factura_m: Literal["REVISAR", "COMPUTAR_CREDITO", "NO_COMPUTAR"] = "REVISAR"
    empresa_es_agente_retencion: bool = False
    alicuotas_iva: list[AlicuotaIVAConfig] = field(default_factory=lambda: list(ALICUOTAS_IVA_DEFAULT))
    regimenes_retencion: list[RegimenRetencion] = field(default_factory=list)


def redondear_moneda(valor: Decimal) -> Decimal:
    return valor.quantize(DOS_DECIMALES, rounding=MODO_REDONDEO)


def formatear_moneda_ar(valor: Decimal) -> str:
    v = redondear_moneda(valor)
    signo = "-" if v < 0 else ""
    s = f"{abs(v):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"{signo}{s}"


def formatear_fecha_ar(fecha: Optional[date]) -> str:
    return fecha.strftime("%d/%m/%Y") if fecha else ""


def periodo_fiscal_mensual(fecha: date) -> str:
    return fecha.strftime("%Y-%m")
